VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Private MUESTRA_ID As Long
Private TIPO_ENSAYO_ID As Long
Private NUMERO_RECEPCION As Long
Private DESIGNACION As String
Private CANTIDAD As Long
Private FECHA_PROCESADO_PIEZAS As String
Private DURACION_FECHA_DESDE As String
Private DURACION_HORA_DESDE As String
Private DURACION_FECHA_HASTA As String
Private DURACION_HORA_HASTA As String
Private ESPESOR As String
Private LOTE_PROBETA_ID As Long
Private MAQUINA As String
Private Incluye_Imagenes As Integer
Private identificacion_laboratorio As String
Private CONDICIONES_AMBIENTALES As String
Private TEMPERATURA As String
Private HUMEDAD As String
Private MATERIAL As String
Private DIMENSION As String
Private REACTIVOS As String
Private REACTIVOS_PROPIOS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setNUMERO_RECEPCION(ByVal dato As Long)
        NUMERO_RECEPCION = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Long)
        CANTIDAD = dato
End Property
Public Property Let setFECHA_PROCESADO_PIEZAS(ByVal dato As String)
        FECHA_PROCESADO_PIEZAS = dato
End Property
Public Property Let setDURACION_FECHA_DESDE(ByVal dato As String)
        DURACION_FECHA_DESDE = dato
End Property
Public Property Let setDURACION_HORA_DESDE(ByVal dato As String)
        DURACION_HORA_DESDE = dato
End Property
Public Property Let setDURACION_FECHA_HASTA(ByVal dato As String)
        DURACION_FECHA_HASTA = dato
End Property
Public Property Let setDURACION_HORA_HASTA(ByVal dato As String)
        DURACION_HORA_HASTA = dato
End Property
Public Property Let setESPESOR(ByVal dato As String)
        ESPESOR = dato
End Property
Public Property Let setLOTE_PROBETA_ID(ByVal dato As Long)
        LOTE_PROBETA_ID = dato
End Property
Public Property Let setMAQUINA(ByVal dato As String)
        MAQUINA = dato
End Property
Public Property Let setINCLUYE_IMAGENES(ByVal dato As Integer)
        Incluye_Imagenes = dato
End Property
Public Property Let setIDENTIFICACION_LABORATORIO(ByVal dato As String)
        identificacion_laboratorio = dato
End Property
Public Property Let setCONDICIONES_AMBIENTALES(ByVal dato As String)
        CONDICIONES_AMBIENTALES = dato
End Property
Public Property Let setTEMPERATURA(ByVal dato As String)
        TEMPERATURA = dato
End Property
Public Property Let setHUMEDAD(ByVal dato As String)
        HUMEDAD = dato
End Property
Public Property Let setMATERIAL(ByVal dato As String)
        MATERIAL = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setREACTIVOS(ByVal dato As String)
        REACTIVOS = dato
End Property
Public Property Let setREACTIVOS_PROPIOS(ByVal dato As String)
        REACTIVOS_PROPIOS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getNUMERO_RECEPCION() As Long
        getNUMERO_RECEPCION = NUMERO_RECEPCION
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getCANTIDAD() As Long
        getCANTIDAD = CANTIDAD
End Property
Public Property Get getFECHA_PROCESADO_PIEZAS() As String
        getFECHA_PROCESADO_PIEZAS = FECHA_PROCESADO_PIEZAS
End Property
Public Property Get getDURACION_FECHA_DESDE() As String
        getDURACION_FECHA_DESDE = DURACION_FECHA_DESDE
End Property
Public Property Get getDURACION_HORA_DESDE() As String
        getDURACION_HORA_DESDE = DURACION_HORA_DESDE
End Property
Public Property Get getDURACION_FECHA_HASTA() As String
        getDURACION_FECHA_HASTA = DURACION_FECHA_HASTA
End Property
Public Property Get getDURACION_HORA_HASTA() As String
        getDURACION_HORA_HASTA = DURACION_HORA_HASTA
End Property
Public Property Get getESPESOR() As String
        getESPESOR = ESPESOR
End Property
Public Property Get getLOTE_PROBETA_ID() As Long
        getLOTE_PROBETA_ID = LOTE_PROBETA_ID
End Property
Public Property Get getMAQUINA() As String
        getMAQUINA = MAQUINA
End Property
Public Property Get getINCLUYE_IMAGENES() As Integer
        getINCLUYE_IMAGENES = Incluye_Imagenes
End Property
Public Property Get getIDENTIFICACION_LABORATORIO() As String
        getIDENTIFICACION_LABORATORIO = identificacion_laboratorio
End Property
Public Property Get getCONDICIONES_AMBIENTALES() As String
        getCONDICIONES_AMBIENTALES = CONDICIONES_AMBIENTALES
End Property
Public Property Get getTEMPERATURA() As String
        getTEMPERATURA = TEMPERATURA
End Property
Public Property Get getHUMEDAD() As String
        getHUMEDAD = HUMEDAD
End Property
Public Property Get getMATERIAL() As String
        getMATERIAL = MATERIAL
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getREACTIVOS() As String
        getREACTIVOS = REACTIVOS
End Property
Public Property Get getREACTIVOS_PROPIOS() As String
        getREACTIVOS_PROPIOS = REACTIVOS_PROPIOS
End Property
Public Function Calcular_Numero_Recepcion()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(Numero_Recepcion) FROM CE_RECEPCION"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                NUMERO_RECEPCION = 1
        Else
                NUMERO_RECEPCION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Public Function carga(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_RECEPCION  WHERE MUESTRA_ID = " & MUESTRA
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                NUMERO_RECEPCION = rs("NUMERO_RECEPCION")
                DESIGNACION = rs("DESIGNACION")
                CANTIDAD = rs("CANTIDAD")
                If Not IsNull(rs("FECHA_PROCESADO_PIEZAS")) Then
                    FECHA_PROCESADO_PIEZAS = rs("FECHA_PROCESADO_PIEZAS")
                Else
                    FECHA_PROCESADO_PIEZAS = ""
                End If
                DURACION_FECHA_DESDE = rs("DURACION_FECHA_DESDE")
                DURACION_HORA_DESDE = rs("DURACION_HORA_DESDE")
                DURACION_FECHA_HASTA = rs("DURACION_FECHA_HASTA")
                DURACION_HORA_HASTA = rs("DURACION_HORA_HASTA")
                ESPESOR = rs("ESPESOR")
                LOTE_PROBETA_ID = rs("LOTE_PROBETA_ID")
                MAQUINA = rs("MAQUINA")
                Incluye_Imagenes = rs("INCLUYE_IMAGENES")
                REACTIVOS = rs("REACTIVOS")
                REACTIVOS_PROPIOS = rs("REACTIVOS_PROPIOS")
                identificacion_laboratorio = rs("IDENTIFICACION_LABORATORIO")
                CONDICIONES_AMBIENTALES = rs("CONDICIONES_AMBIENTALES")
                TEMPERATURA = IIf(IsNull(rs("TEMPERATURA")), "", rs("TEMPERATURA"))
                HUMEDAD = IIf(IsNull(rs("HUMEDAD")), "", rs("HUMEDAD"))
                MATERIAL = rs("MATERIAL")
                DIMENSION = rs("DIMENSION")
                
                carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        carga = False
        MsgBox "Error al cargar los datos(clsCE_RECEPCION)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        ' JGM : FALTA EL ENVIADO_PAQUETE
        consulta = "INSERT INTO CE_RECEPCION  " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & TIPO_ENSAYO_ID & "," & NUMERO_RECEPCION & ",'" & DESIGNACION & "'," & _
                        CANTIDAD & "," & FECHA_PROCESADO_PIEZAS & "," & "'" & DURACION_FECHA_DESDE & "'" & "," & _
                        "'" & DURACION_HORA_DESDE & "'" & "," & "'" & DURACION_FECHA_HASTA & "'" & "," & "'" & DURACION_HORA_HASTA & "'" & "," & _
                        "'" & ESPESOR & "'" & "," & LOTE_PROBETA_ID & "," & "'" & MAQUINA & "'," & Incluye_Imagenes & "," & identificacion_laboratorio & ",'" & CONDICIONES_AMBIENTALES & "'," & _
                        IIf(TEMPERATURA = "", "null", TEMPERATURA) & "," & _
                        IIf(HUMEDAD = "", "null", HUMEDAD) & "," & _
                        "'" & MATERIAL & "','" & DIMENSION & "','" & REACTIVOS & "','" & REACTIVOS_PROPIOS & "',0" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCE_RECEPCION)", vbCritical, Err.Description
End Function
Public Function InformarTipoEnsayo(idMuestra As Long, idTipoEnsayo As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCION  SET " & _
                   " TIPO_ENSAYO_ID = " & idTipoEnsayo & _
                   " WHERE MUESTRA_ID = " & idMuestra
        execute_bd consulta
        InformarTipoEnsayo = True
        Exit Function
fallo:
        InformarTipoEnsayo = False
        MsgBox "Error al InformarTipoEnsayo (clsCE_RECEPCION)", vbCritical, Err.Description
End Function

Public Function Informar_cantidad(MUESTRA As Long, CANTIDAD As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCION  SET " & _
                   " CANTIDAD = " & CANTIDAD & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Informar_cantidad = True
        Exit Function
fallo:
        Informar_cantidad = False
        MsgBox "Error al Informar_Cantidad (clsCE_RECEPCION)", vbCritical, Err.Description
End Function
Public Function Informar_Identificacion_Laboratorio(MUESTRA As Long, Identificacion As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCION  SET " & _
                   " Identificacion_Laboratorio = " & Identificacion & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Informar_Identificacion_Laboratorio = True
        Exit Function
fallo:
        Informar_Identificacion_Laboratorio = False
        MsgBox "Error al Informar_Identificacion_Laboratorio (clsCE_RECEPCION)", vbCritical, Err.Description
End Function
'Public Function Informar_Condiciones_Ambientales(MUESTRA As Long, CA As String) As Boolean
'        On Error GoTo fallo
'        Dim consulta As String
'        consulta = "UPDATE CE_RECEPCION  SET " & _
'                   " CONDICIONES_AMBIENTALES = '" & CA & "'" & _
'                   " WHERE MUESTRA_ID = " & MUESTRA
'        execute_bd consulta
'        Informar_Condiciones_Ambientales = True
'        Exit Function
'fallo:
'        Informar_Condiciones_Ambientales = False
'        MsgBox "Error al Informar_Condiciones_Ambientales (clsCE_RECEPCION)", vbCritical, Err.Description
'End Function
Public Function Informar_Duracion_Ensayo(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCION  " & _
                   "  SET DURACION_FECHA_DESDE = '" & DURACION_FECHA_DESDE & "'," & _
                   "      DURACION_HORA_DESDE  = '" & DURACION_HORA_DESDE & "'," & _
                   "      DURACION_FECHA_HASTA = '" & DURACION_FECHA_HASTA & "'," & _
                   "      DURACION_HORA_HASTA  = '" & DURACION_HORA_HASTA & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Informar_Duracion_Ensayo = True
        Exit Function
fallo:
        Informar_Duracion_Ensayo = False
        MsgBox "Error al Informar_Duracion_Ensayo (clsCE_RECEPCION)", vbCritical, Err.Description
End Function
Public Function Informar_registro(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        'M1104 : Fecha procesado a tipo Date (Eliminacion de las comillas)
        consulta = "UPDATE CE_RECEPCION  " & _
                   "  SET DURACION_FECHA_DESDE = '" & DURACION_FECHA_DESDE & "'," & _
                   "      DURACION_HORA_DESDE  = '" & DURACION_HORA_DESDE & "'," & _
                   "      DURACION_FECHA_HASTA = '" & DURACION_FECHA_HASTA & "'," & _
                   "      FECHA_PROCESADO_PIEZAS  = " & FECHA_PROCESADO_PIEZAS & "," & _
                   "      LOTE_PROBETA_ID = " & LOTE_PROBETA_ID & "," & _
                   "      ESPESOR = '" & ESPESOR & "'," & _
                   "      MAQUINA = '" & MAQUINA & "'," & _
                   "      REACTIVOS = '" & REACTIVOS & "'," & _
                   "      REACTIVOS_PROPIOS = '" & REACTIVOS_PROPIOS & "'," & _
                   "      CONDICIONES_AMBIENTALES = '" & CONDICIONES_AMBIENTALES & "'," & _
                   "      MATERIAL = '" & MATERIAL & "'," & _
                   "      DIMENSION = '" & DIMENSION & "'," & _
                   "      TEMPERATURA = " & TEMPERATURA & "," & _
                   "      HUMEDAD = " & HUMEDAD & "," & _
                   "      DURACION_HORA_HASTA  = '" & DURACION_HORA_HASTA & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Informar_registro = True
        Exit Function
fallo:
        Informar_registro = False
        MsgBox "Error al Informar_registro (clsCE_RECEPCION)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_RECEPCION  " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCE_RECEPCION)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_RECEPCION  ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_recepcion(RECEPCION As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_RECEPCION  " & _
                   " WHERE NUMERO_RECEPCION = " & RECEPCION & _
                   " ORDER BY MUESTRA_ID"
        Set Listado_por_recepcion = datos_bd(consulta)
End Function
Public Function Listado_ids_por_recepcion(NUMERO_RECEPCION As Long) As String
        Dim consulta As String
        Dim rs As ADODB.Recordset
   On Error GoTo Listado_ids_por_recepcion_Error

        consulta = "SELECT group_concat(MUESTRA_ID) FROM CE_RECEPCION WHERE NUMERO_RECEPCION = " & NUMERO_RECEPCION
        Set rs = datos_bd(consulta)
        Dim s As String
        If rs.RecordCount > 0 Then
            s = rs(0)
        Else
            s = "SELECT MUESTRA_ID FROM CE_RECEPCION WHERE NUMERO_RECEPCION = " & NUMERO_RECEPCION
        End If
        Listado_ids_por_recepcion = s

   On Error GoTo 0
   Exit Function

Listado_ids_por_recepcion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Listado_ids_por_recepcion of Módulo de clase clsCe_recepcion"
End Function
Public Function Listado_NoIniciadas() As ADODB.Recordset
        Dim consulta As String
        Dim DIAS As String
        Dim opar As New clsParametros
        opar.carga parametros.DIAS_AVISO_CONTROLES_NOINICIADOS, ""
        If opar.getVALOR = "" Then
            DIAS = 30
        Else
            DIAS = opar.getVALOR
        End If
'        consulta = " select b.ID_MUESTRA,b.ID_GENERAL,c.NOMBRE, b.REFERENCIA_CLIENTE, b.FECHA_RECEPCION f1,date(date_sub(current_date,interval " & dias & " day)) as f2 " & _
'                   "   from ce_recepcion a inner join muestras b on a.MUESTRA_ID = b.ID_MUESTRA " & _
'                   "                    inner join clientes c on b.CLIENTE_ID = c.ID_CLIENTE " & _
'                   " where a.DURACION_FECHA_DESDE = '' or a.DURACION_FECHA_DESDE = '01-01-1900' " & _
'                   "   and b.CERRADA = 0 and b.ANALISIS_MODIFICADO = 2 " & _
'                   " Having f1 > f2 " & _
'                   " order by f1, id_muestra asc"
        consulta = " SELECT b.ID_MUESTRA,b.ID_GENERAL,c.NOMBRE, b.REFERENCIA_CLIENTE, b.FECHA_RECEPCION f1, DATE(DATE_SUB(CURRENT_DATE, INTERVAL " & DIAS & " DAY)) AS f2, " & _
                   "        d.DURACION_FECHA_DESDE,d.DURACION_FECHA_HASTA " & _
                   "   FROM muestras b " & _
                   "  INNER JOIN clientes c ON b.CLIENTE_ID = c.ID_CLIENTE " & _
                   "   LEFT JOIN ce_recepcion d ON b.ID_MUESTRA = d.MUESTRA_ID " & _
                   "  where b.CERRADA <> 1 and b.cerrada <> 3 And b.ANULADA = 0 and b.id_muestra <> 0 " & _
                   " Having f1 < f2 " & _
                   " ORDER BY f1, id_muestra ASC  "
        Set Listado_NoIniciadas = datos_bd(consulta)
End Function

Public Function Incluir_Imagenes(MUESTRA As Long, imagenes As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCION  SET " & _
                   " INCLUYE_IMAGENES = " & imagenes & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Incluir_Imagenes = True
        Exit Function
fallo:
        Incluir_Imagenes = False
        MsgBox "Error al Incluir_Imagenes (clsCE_RECEPCION)", vbCritical, Err.Description
End Function
'M0957-I
Public Function ListadoSubcontratables(stranno As String, strCODIGO As String, strRefCliente As String, strEnsayo As String, strContrata As String, strEnviado As String) As ADODB.Recordset
    On Error GoTo fallo
    
    Dim consulta As String
    
'    consulta = "SELECT CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) CODIGO      , M.REFERENCIA_CLIENTE" & _
'               "     , TA.NOMBRE, TE.NOMBRE, TE.PROCESO, TEC.PRECIO, 0      , M.ID_MUESTRA," & _
'               "     TE.ID_TIPO_ENSAYO ID_TE, R.NUMERO_RECEPCION, R.ENVIADO_PAQUETE  " & _
'               "  FROM MUESTRAS M, TIPOS_MUESTRA TM, TIPOS_ANALISIS TA " & _
'               "     , CE_RECEPCION R, CE_TIPOS_ENSAYOS TE, CE_TIPOS_ENSAYOS_CONTRATAS TEC" & _
'               " WHERE M.ID_MUESTRA = R.MUESTRA_ID " & _
'               "   AND M.TIPO_MUESTRA_ID =   TM.ID_TIPO_MUESTRA " & _
'               "   AND R.TIPO_ENSAYO_ID =    TE.ID_TIPO_ENSAYO " & _
'               "   AND R.TIPO_ENSAYO_ID =    TEC.TIPO_ENSAYO_ID " & _
'               "   AND TE.TIPO_ANALISIS_ID = TA.ID_TIPO_ANALISIS " & _
'               "   AND M.CERRADA <> 1 " & _
'               "   AND M.ANULADA = 0 " & _
'               "   AND TE.ES_SUBCONTRATABLE = 1 " & _
'               "   AND R.ENVIADO_PAQUETE IN (" & strEnviado & ")" & _
'               "   AND M.ANNO LIKE " & stranno & _
'               "   AND CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) LIKE '%" & strCODIGO & "%' " & _
'               "   AND M.REFERENCIA_CLIENTE LIKE '%" & strRefCliente & "%' " & _
'               "   AND TE.NOMBRE LIKE '%" & strEnsayo & "%' " & _
'               " ORDER BY CODIGO "
    
    consulta = "SELECT CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) CODIGO      , M.REFERENCIA_CLIENTE" & _
               "     , TA.NOMBRE, TE.NOMBRE, PB.NOMBRE AS PROCESO, TA.PRECIO, 0      , M.ID_MUESTRA," & _
               "     TE.ID_TIPO_ENSAYO ID_TE, R.NUMERO_RECEPCION, R.ENVIADO_PAQUETE  " & _
               "  FROM MUESTRAS M, TIPOS_MUESTRA TM, TIPOS_ANALISIS TA " & _
               "     , CE_RECEPCION R, CE_TIPOS_ENSAYOS TE, PROCESOS_BASE PB " & _
               " WHERE M.ID_MUESTRA = R.MUESTRA_ID " & _
               "   AND M.TIPO_MUESTRA_ID =   TM.ID_TIPO_MUESTRA " & _
               "   AND R.TIPO_ENSAYO_ID =    TE.ID_TIPO_ENSAYO " & _
               "   AND TE.TIPO_ANALISIS_ID = TA.ID_TIPO_ANALISIS " & _
               "   AND TE.PROCESO_BASE_ID = PB.ID_PROCESO_BASE " & _
               "   AND M.CERRADA <> 1 " & _
               "   AND M.ANULADA = 0 " & _
               "   AND TE.ES_SUBCONTRATABLE = 1 " & _
               "   AND R.ENVIADO_PAQUETE IN (" & strEnviado & ")" & _
               "   AND M.ANNO LIKE " & stranno & _
               "   AND CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) LIKE '%" & strCODIGO & "%' " & _
               "   AND M.REFERENCIA_CLIENTE LIKE '%" & strRefCliente & "%' " & _
               "   AND TE.NOMBRE LIKE '%" & strEnsayo & "%' " & _
               " ORDER BY CODIGO "
    Set ListadoSubcontratables = datos_bd(consulta)
    Exit Function
    
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function

Public Function marcar_ensayo_enviado_paquete(lngMuestra_ID As Long, lngEnsayo_ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    consulta = "UPDATE CE_RECEPCION " & _
               "   SET ENVIADO_PAQUETE = 1 " & _
               " WHERE MUESTRA_ID = " & lngMuestra_ID & _
               "   AND TIPO_ENSAYO_ID = " & lngEnsayo_ID
    execute_bd consulta
    marcar_ensayo_enviado_paquete = True
    Exit Function
    
fallo:
    marcar_ensayo_enviado_paquete = False
    MsgBox "Error al marcar el ensayo como enviado en paquete.", vbCritical, Err.Description
End Function
'M0957-F
Public Function desmarcar_ensayo_enviado_paquete(lngMuestra_ID As Long, lngEnsayo_ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    consulta = "UPDATE CE_RECEPCION " & _
               "   SET ENVIADO_PAQUETE = 0 " & _
               " WHERE MUESTRA_ID = " & lngMuestra_ID & _
               "   AND TIPO_ENSAYO_ID = " & lngEnsayo_ID
    execute_bd consulta
    desmarcar_ensayo_enviado_paquete = True
    Exit Function
    
fallo:
    desmarcar_ensayo_enviado_paquete = False
    MsgBox "Error al desmarcar el ensayo como enviado en paquete.", vbCritical, Err.Description
End Function

