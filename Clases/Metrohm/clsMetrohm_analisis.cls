VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMetrohm_analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMETROHM_ANALISIS
'***************************************************************
Private ID As Long
Private MUESTRA_ID As Long
Private TIPO_DETERMINACION_ID As Long
Private resultado As String
Private ES_DUPLICADA As Long
Private fecha As String
Private procesada As Long
Private USUARIO_ID As Long
Private TS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMETROHM_ANALISIS
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        resultado = dato
End Property
Public Property Let setES_DUPLICADA(ByVal dato As Long)
        ES_DUPLICADA = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setPROCESADA(ByVal dato As Long)
        procesada = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMETROHM_ANALISIS
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = resultado
End Property
Public Property Get getES_DUPLICADA() As Long
        getES_DUPLICADA = ES_DUPLICADA
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getPROCESADA() As Long
        getPROCESADA = procesada
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMETROHM_ANALISIS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM METROHM_ANALISIS WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                MUESTRA_ID = rs("MUESTRA_ID")
                TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                resultado = rs("RESULTADO")
                ES_DUPLICADA = rs("ES_DUPLICADA")
                fecha = rs("FECHA")
                procesada = rs("PROCESADA")
                USUARIO_ID = rs("USUARIO_ID")
                TS = rs("TS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMetrohm_analisis)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO METROHM_ANALISIS " & _
                   "  VALUES (" & _
                        ID & "," & MUESTRA_ID & "," & TIPO_DETERMINACION_ID & "," & _
                        "'" & resultado & "'" & "," & ES_DUPLICADA & "," & "'" & fecha & "'" & ",0,null,null" & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMetrohm_analisis)", vbCritical, Err.Description
End Function
Public Function marcarProcesado(ID As Long, USUARIO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE METROHM_ANALISIS SET " & _
                        " PROCESADA = 1," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TS = current_timestamp " & _
                " WHERE ID = " & ID
        execute_bd consulta
        marcarProcesado = True
        Exit Function
fallo:
        marcarProcesado = False
        MsgBox "Error al marcarProcesado (clsMetrohm_analisis)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM METROHM_ANALISIS " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMetrohm_analisis)", vbCritical, Err.Description
End Function
Public Function Listado(procesada As Boolean, fdesde As String, fhasta As String, deter As String, muestra As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1 and b.CERRADA = 0 "
        If procesada Then
            filtro = filtro & " AND a.procesada = 0 "
        Else
            filtro = filtro & " AND a.procesada = 1 "
        End If
        filtro = filtro & " and date_format(a.fecha,'%Y-%m-%d') >= '" & Format(fdesde, "yyyy-mm-dd") & "'"
        filtro = filtro & " and date_format(a.fecha,'%Y-%m-%d') <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
        If deter <> "" Then
            filtro = filtro & " and d.nombre like '%" & deter & "%'"
        End If
        If muestra <> "" Then
            filtro = filtro & " and concat(c.CODIGO,'-',LPAD(b.ID_PARTICULAR,4,'0')) like '%" & muestra & "%'"
        End If
        consulta = "select a.ID, a.MUESTRA_ID,a.TIPO_DETERMINACION_ID " & _
                    " ,e.ID_DETERMINACION,f.FORMULA_ID,f.ID_CAMPO,a.ES_DUPLICADA  " & _
                    " ,b.ID_GENERAL, concat(c.CODIGO,'-',LPAD(b.ID_PARTICULAR,4,'0')) as ID_PARTICULAR, concat(d.NOMBRE,' (',a.ES_DUPLICADA + 1,')') as DETERMINACION,f.NOMBRE AS CAMPO, a.RESULTADO, h.NOMBRE as UNIDAD,a.FECHA " & _
                    " ,IFNULL(g.MINIMO,'') AS MINIMO,IFNULL(g.MAXIMO,'') AS MAXIMO, IFNULL(ec.DIF_AVISO,'') AS DIF_AVISO, IFNULL(ec.INCERTIDUMBRE,'') AS INCERTIDUMBRE " & _
                    " from metrohm_analisis a  " & _
                    " inner join muestras b on a.MUESTRA_ID = b.ID_MUESTRA " & _
                    " inner join tipos_muestra c on b.TIPO_MUESTRA_ID = c.ID_TIPO_MUESTRA " & _
                    " inner join tipos_determinacion d on a.TIPO_DETERMINACION_ID = d.ID_TIPO_DETERMINACION " & _
                    " inner join determinaciones e on b.ID_MUESTRA = e.MUESTRA_ID and e.TIPO_DETERMINACION_ID = d.ID_TIPO_DETERMINACION " & _
                    " inner join determinaciones_conf ec on e.ID_DETERMINACION = ec.DETERMINACION_ID " & _
                    " inner join formulas_campos f on e.FORMULA_ID = f.FORMULA_ID and f.ES_SOLUCION = 1 " & _
                    " left join determinaciones_analisis g on IF(b.BANO_ID <> 0, b.BANO_ID = g.BANO_ID, b.TIPO_ANALISIS_ID = g.TIPO_ANALISIS_ID) and a.TIPO_DETERMINACION_ID = g.TIPO_DETERMINACION_ID " & _
                    " left join unidades h on f.UNIDAD_ID = h.ID_UNIDAD " & _
                    filtro & _
                    " order by a.MUESTRA_ID, a.TIPO_DETERMINACION_ID, a.ES_DUPLICADA, a.FECHA "
        Set Listado = datos_bd(consulta)
End Function
'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMETROHM_ANALISIS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

