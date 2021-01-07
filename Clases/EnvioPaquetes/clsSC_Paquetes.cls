VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSC_Paquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSC_PAQUETES
'***************************************************************
Private ID_PAQUETE As Long
Private CENTRO_ID As Integer
Private CODIGO_SC As String
Private PRESUPUESTO As String
Private OBSERVACIONES As String
Private SUBCONTRATA_ID As Long
Private CLIENTE_ID As Long
Private fecha_creacion As String
Private HORA_CREACION As String
'M0959-I
Private FACTURA_RECIBIDA As Integer
'M0957-I
'Private NFactura As Long
Private NFACTURA As String
'M0957-F
Private ffactura As String
'M0959-F

Private USUARIO_ID As Long
'M0957-I
Private APROBADOR_ID As Long
Private FECHA_APROBACION As String
Private RECEPTOR_ID As Long
Private FECHA_RECEPCION As String
Private ESTADO As Long
'M0957-F
'M0957-I
Private tipo As Long
'M0957-F
'M1257-I
Private subtipo As Long
'M1257-F
'M1274-I
Private CUSERID As String
Private FTIMESTP As String
Private EDICION As Long
Private moneda As Integer
'M1274-F

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSC_PAQUETES
'***************************************************************
Public Property Let setID_PAQUETE(ByVal dato As Long)
        ID_PAQUETE = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Integer)
        CENTRO_ID = dato
End Property
Public Property Let setCODIGO_SC(ByVal dato As String)
        CODIGO_SC = dato
End Property
Public Property Let setPRESUPUESTO(ByVal dato As String)
        PRESUPUESTO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setSUBCONTRATA_ID(ByVal dato As Long)
        SUBCONTRATA_ID = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setFECHA_CREACION(ByVal dato As String)
        fecha_creacion = dato
End Property
Public Property Let setHORA_CREACION(ByVal dato As String)
        HORA_CREACION = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
'M0959-I
Public Property Let setFACTURA_RECIBIDA(ByVal dato As Long)
        FACTURA_RECIBIDA = dato
End Property
Public Property Let setNFACTURA(ByVal dato As String)
        NFACTURA = dato
End Property
Public Property Let setFFACTURA(ByVal dato As String)
        ffactura = dato
End Property
'M0959-F
'M0957-I
Public Property Let setAPROBADOR_ID(ByVal dato As Long)
        APROBADOR_ID = dato
End Property
Public Property Let setFECHA_APROBACION(ByVal dato As String)
        FECHA_APROBACION = dato
End Property
Public Property Let setRECEPTOR_ID(ByVal dato As Long)
        RECEPTOR_ID = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        FECHA_RECEPCION = dato
End Property
Public Property Let setESTADO(ByVal dato As Long)
        ESTADO = dato
End Property
'M0957-F
'M0957-I
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
'M0957-F
'M1257-I
Public Property Let setSUBTIPO(ByVal dato As Long)
        subtipo = dato
End Property
'M1257-F
'M1274-I
Public Property Let setEDICION(ByVal dato As Long)
        EDICION = dato
End Property
'M1274-F
Public Property Let setMONEDA(ByVal dato As Integer)
        moneda = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSC_PAQUETES
'***************************************************************
Public Property Get getID_PAQUETE() As Long
        getID_PAQUETE = ID_PAQUETE
End Property
Public Property Get getCENTRO_ID() As Integer
        getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getCODIGO_SC() As String
        getCODIGO_SC = CODIGO_SC
End Property
Public Property Get getPRESUPUESTO() As String
        getPRESUPUESTO = PRESUPUESTO
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getSUBCONTRATA_ID() As Long
        getSUBCONTRATA_ID = SUBCONTRATA_ID
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getFECHA_CREACION() As String
        getFECHA_CREACION = fecha_creacion
End Property
Public Property Get getHORA_CREACION() As String
        getHORA_CREACION = HORA_CREACION
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
'M0959-I
Public Property Get getFACTURA_RECIBIDA() As Long
        getFACTURA_RECIBIDA = FACTURA_RECIBIDA
End Property
Public Property Get getNFACTURA() As String
        getNFACTURA = NFACTURA
End Property
Public Property Get getFFACTURA() As String
        getFFACTURA = ffactura
End Property
'M0959-F
'M0957-I
Public Property Get getAPROBADOR_ID() As Long
        getAPROBADOR_ID = APROBADOR_ID
End Property
Public Property Get getFECHA_APROBACION() As String
        getFECHA_APROBACION = FECHA_APROBACION
End Property
Public Property Get getRECEPTOR_ID() As Long
        getRECEPTOR_ID = RECEPTOR_ID
End Property
Public Property Get getFECHA_RECEPCION() As String
        getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getESTADO() As Long
        getESTADO = ESTADO
End Property
'M0957-F
'M0957-I
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
'M0957-F
'M1257-I
Public Property Get getSUBTIPO() As Long
        getSUBTIPO = subtipo
End Property
'M1257-F
'M1254-I
Public Property Get getEDICION() As Long
        getEDICION = EDICION
End Property
'M1254-F
Public Property Get getMONEDA() As Integer
        getMONEDA = moneda
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSC_PAQUETES
'***************************************************************
Public Function Carga(ID As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
'M1274-I
'        consulta = "SELECT * FROM SC_PAQUETES WHERE ID_PAQUETE = " & ID
        consulta = "SELECT * FROM SC_PAQUETES WHERE ID_PAQUETE = " & ID & " AND EDICION = " & EDICION
'M1274-F
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PAQUETE = rs("ID_PAQUETE")
                CENTRO_ID = rs("CENTRO_ID")
                CODIGO_SC = rs("CODIGO_SC")
                PRESUPUESTO = rs("PRESUPUESTO")
                OBSERVACIONES = rs("OBSERVACIONES")
                SUBCONTRATA_ID = rs("SUBCONTRATA_ID")
                CLIENTE_ID = rs("CLIENTE_ID")
                fecha_creacion = Format(rs("FECHA_CREACION"), "yyyy-mm-dd")
                HORA_CREACION = rs("HORA_CREACION")
                USUARIO_ID = rs("USUARIO_ID")
                'M0959-I
                FACTURA_RECIBIDA = rs("FACTURA_RECIBIDA")
                NFACTURA = rs("NFACTURA")
                ffactura = Format(rs("FFACTURA"), "yyyy-mm-dd")
                'M0959-F
                'M0957-I
                APROBADOR_ID = rs("APROBADOR_ID")
                FECHA_APROBACION = Format(rs("FECHA_APROBACION"), "yyyy-mm-dd")
                RECEPTOR_ID = rs("RECEPTOR_ID")
                FECHA_RECEPCION = Format(rs("FECHA_RECEPCION"), "yyyy-mm-dd")
                ESTADO = rs("ESTADO")
                'M0957-F
                'M0957-I
                tipo = rs("TIPO")
                'M0957-F
                'M1257-I
                subtipo = rs("SUBTIPO")
                'M1257-F
                'M1274-F
                EDICION = rs("EDICION")
                'M1274-F
                moneda = rs("MONEDA")
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSC_Paquetes)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_PAQUETE) FROM SC_PAQUETES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PAQUETE = 1
        Else
                ID_PAQUETE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function CrearCodigoSC() As String
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select max(codigo_sc)  from sc_paquetes" & _
               " where codigo_sc like '%/" & Format(Date, "yy") & "%'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            CODIGO_SC = "0001" & "/" & Right(Year(Date), 2)
    Else
            CODIGO_SC = Format((Left(rs.Fields(0), 4) + 1), "0000")
            CODIGO_SC = CODIGO_SC & "/" & Right(Year(Date), 2)
    End If
    Set rs = Nothing
'M1274-F
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_PAQUETE = 0 Then
            Me.CrearID
            Me.CrearCodigoSC
'        Else
'            Me.CrearId
        End If
        
        consulta = "INSERT INTO SC_PAQUETES (" & _
                    "      ID_PAQUETE, CENTRO_ID,CODIGO_SC, PRESUPUESTO " & _
                    "    , OBSERVACIONES, SUBCONTRATA_ID,CLIENTE_ID " & _
                    "    , FECHA_CREACION, HORA_CREACION" & _
                    "    , FACTURA_RECIBIDA, NFACTURA, FFACTURA" & _
                    "    , USUARIO_ID, APROBADOR_ID, FECHA_APROBACION, ESTADO, RECEPTOR_ID, FECHA_RECEPCION " & _
                    "    , TIPO, SUBTIPO" & _
                    "    , CUSERID, FTIMESTP, EDICION, MONEDA " & _
                    ") VALUES (" & _
                           ID_PAQUETE & "," & CENTRO_ID & ",'" & CODIGO_SC & "'" & ", " & _
                           PRESUPUESTO & "," & "'" & OBSERVACIONES & "'" & ", " & _
                           SUBCONTRATA_ID & "," & CLIENTE_ID & "," & "'" & Format(fecha_creacion, "yyyy-mm-dd") & "'" & ", " & "'" & HORA_CREACION & "'" & ", " & _
                           FACTURA_RECIBIDA & ", '" & NFACTURA & "'," & "'" & ffactura & "'," & _
                           USUARIO_ID & ", " & APROBADOR_ID & ",'" & FECHA_APROBACION & "'," & ESTADO & "," & _
                           RECEPTOR_ID & ",'" & FECHA_RECEPCION & "'," & _
                           tipo & "," & subtipo & "," & _
                           "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP" & "," & EDICION & "," & moneda & ")"
        
        execute_bd consulta
'M1274-I
        If EDICION > 1 Then
           consulta = "UPDATE SC_PAQUETES" & _
                      " SET ESTADO = " & SC_ESTADO_HISTORICO & _
                      " WHERE ID_PAQUETE = " & ID_PAQUETE & _
                      " AND EDICION < " & EDICION
           execute_bd consulta
        End If
'M1274-F
        Insertar = ID_PAQUETE
        Exit Function
        
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSC_Paquetes)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE SC_PAQUETES SET " & _
                   "       PRESUPUESTO = '" & PRESUPUESTO & "'" & _
                   "     , SUBCONTRATA_ID = " & SUBCONTRATA_ID & _
                   "     , CLIENTE_ID = " & CLIENTE_ID & _
                   "     , OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                   "     , CUSERID = '" & USUARIO.getUSUARIO & "'" & _
                   "     , NFACTURA = '" & NFACTURA & "'" & _
                   "     , FFACTURA = '" & ffactura & "'" & _
                   "     , FACTURA_RECIBIDA = " & FACTURA_RECIBIDA & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP" & _
                   "     , EDICION = " & EDICION & _
                   "     , SUBTIPO = " & subtipo & _
                   "     , CENTRO_ID = " & CENTRO_ID & _
                   "     , MONEDA = " & moneda & _
                   " WHERE ID_PAQUETE = " & ID & _
                   " AND EDICION = " & EDICION
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSC_Paquetes)", vbCritical, Err.Description
End Function

'M0957-I
Public Function Tramitar(ID As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE SC_PAQUETES SET " & _
                   "       ESTADO = " & SC_ESTADO_TRAMITADO & _
                   "     , APROBADOR_ID = " & USUARIO.getID_EMPLEADO & _
                   "     , FECHA_APROBACION = '" & Format(Date, "yyyy-mm-dd") & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE ID_PAQUETE = " & ID & _
                   " AND EDICION = " & EDICION
        
        execute_bd consulta
        Tramitar = True
        Exit Function
fallo:
        Tramitar = False
        MsgBox "Error al Tramitar (clsSC_Paquetes)", vbCritical, Err.Description
End Function
'M0957-F
'M1147-I
Public Function finalizar(ID As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE SC_PAQUETES SET " & _
                   "       ESTADO = " & SC_ESTADO_RECIBIDO & _
                   "     , RECEPTOR_ID = " & USUARIO.getID_EMPLEADO & _
                   "     , FECHA_RECEPCION = '" & Format(Date, "yyyy-mm-dd") & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE ID_PAQUETE = " & ID & _
                   " AND EDICION = " & EDICION
        
        execute_bd consulta
        finalizar = True
        Exit Function
fallo:
        finalizar = False
        MsgBox "Error al Finalizar (clsSC_Paquetes)", vbCritical, Err.Description
End Function
'M1147-F
Public Function Eliminar(ID As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        ' se vuelven a poner las determinaciones como sin enviar
        execute_bd "UPDATE DETERMINACIONES D, SC_PAQUETES_DETALLE SCPD " & _
                   "   SET ENVIADO_PAQUETE = 0 " & _
                   " WHERE D.ID_DETERMINACION = SCPD.DETERMINACION_ID " & _
                   "   AND SCPD.PAQUETE_ID = " & ID & _
                   "   AND SCPD.EDICION = " & EDICION
        'CE
        execute_bd "UPDATE CE_RECEPCION CE, SC_PAQUETES_DETALLE SCPD " & _
                   " SET ENVIADO_PAQUETE = 0 " & _
                   " WHERE CE.MUESTRA_ID = SCPD.MUESTRA_ID " & _
                   " AND SCPD.PAQUETE_ID = " & ID & _
                   " AND SCPD.EDICION = " & EDICION
               
'                   " AND CE.TIPO_ENSAYO_ID = SCPD.TIPO_ENSAYO_ID " & _

        'Si existe, recuperamos la versión anterior
        If EDICION > 1 Then
           execute_bd "UPDATE SC_PAQUETES " & _
                      " SET ESTADO = " & SC_ESTADO_TRAMITADO & _
                      " WHERE ID_PAQUETE = " & ID & " AND EDICION = " & (EDICION - 1)
                                          
        End If
        
        ' se quitan las determinaciones del detalle del paquete
        execute_bd "DELETE FROM SC_PAQUETES_DETALLE WHERE PAQUETE_ID = " & ID & " AND EDICION = " & EDICION
        
        ' se quitan los conceptos del detalle del paquete
        execute_bd "DELETE FROM SC_PAQUETES_DETALLE_GENERICO WHERE PAQUETE_ID = " & ID & " AND EDICION = " & EDICION
        
        ' se elimina el paquete
        execute_bd "DELETE FROM SC_PAQUETES WHERE ID_PAQUETE = " & ID & " AND EDICION = " & EDICION
        Eliminar = True
        Exit Function
        
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSC_Paquetes)", vbCritical, Err.Description
End Function
Public Function Listado_parcial(NUM_REGISTROS) As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * " & _
                   " FROM SC_PAQUETES" & _
                   " ORDER BY ID_PAQUETE DESC " & _
                   " LIMIT " & NUM_REGISTROS
        Set Listado_parcial = datos_bd(consulta)
End Function

Public Function Listado(strCodigoSC As String, strUsuario As String, strContrata As String, strFechaDesde As String, strFechaHasta As String, facturaRecibida As String, ffactura As String, FFACTURAf As String, NFACTURA As String, ESTADO As Long, tipo As Long, subtipo As Long, Centro As Integer) As ADODB.Recordset
    Dim consulta As String
    Dim strFactura As String
    Dim facturai As Long
    Dim facturaf As Long
    Dim strTramitado As String
    Dim strTipo As String
    NFACTURA = Trim(NFACTURA)
    If (facturaRecibida <> 0) Then
        strFactura = " AND FACTURA_RECIBIDA = " & facturaRecibida
        If ffactura <> "" Then
            strFactura = strFactura + " AND FFACTURA>='" & ffactura & "' AND FFACTURA <='" & FFACTURAf & "'"
        End If
        strFactura = strFactura & " AND NFACTURA LIKE '%" & NFACTURA & "%'"
    End If
    If ESTADO <> 0 Then
        strTramitado = " AND P.ESTADO = " & ESTADO
    End If
    Dim filtro As String
    If Trim(strCodigoSC) <> "" Then
        filtro = filtro & " AND P.CODIGO_SC LIKE '%" & strCodigoSC & "%' "
    End If
    If Trim(strUsuario) <> "" Then
        filtro = filtro & " AND U.NOMBRE LIKE '%" & strUsuario & "%' "
    End If
    If Trim(strContrata) <> "" Then
        filtro = filtro & " AND PR.NOMBRE = '" & strContrata & "' "
    End If
    If tipo <> 0 Then
       strTipo = " AND P.TIPO = " & tipo
    End If
    If subtipo <> 0 Then
       strTipo = strTipo & " AND P.SUBTIPO = " & subtipo
    End If
    If Centro <> 0 Then
       strTipo = strTipo & " AND P.CENTRO_ID = " & Centro
    End If
    strFactura = strFactura & " AND NFACTURA LIKE '%" & NFACTURA & "' "
    consulta = "SELECT P.CODIGO_SC, PR.NOMBRE CONTRATA, P.PRESUPUESTO  " & _
                ", P.FECHA_CREACION, U.NOMBRE USUARIO , P.ID_PAQUETE, PR.ID_PROVEEDOR, P.ESTADO, P.FECHA_APROBACION, AP.NOMBRE USUARIO, P.TIPO, P.SUBTIPO, temp.suma, P.EDICION, D1.PARAMETROS AS COLOR, D1.DESCRIPCION AS STRESTADO, D2.PARAMETROS, D3.DESCRIPCION AS SUBTIPO,centros.NOMBRE,MON.DESCRIPCION " & _
                " FROM SC_PAQUETES P" & _
                " inner join proveedores PR on p.subcontrata_id = PR.ID_PROVEEDOR" & _
                " inner join usuarios u on P.usuario_id = u.ID_EMPLEADO" & _
                " inner join usuarios AP on P.aprobador_id = AP.ID_EMPLEADO" & _
                " left join decodificadora D1 on P.ESTADO = D1.valor and D1.codigo  = 199" & _
                " left join decodificadora D2 on P.TIPO = D2.valor and D2.codigo  = 500" & _
                " left join decodificadora D3 on P.SUBTIPO = D3.valor and D3.codigo  = 198" & _
                " LEFT JOIN decodificadora MON ON P.MONEDA = MON.VALOR and MON.CODIGO = 180 " & _
                " left join (select tobjeto,cobjeto,proveedor_id, sum(BI) as SUMA from proveedores_facturas WHERE TOBJETO in (31,32,33) GROUP BY tobjeto, cobjeto) as temp on p.id_paquete=temp.cobjeto and p.subcontrata_id = temp.proveedor_id" & _
                " left join centros as centros on p.centro_id = centros.id_centro " & _
                " WHERE P.FECHA_CREACION >= '" & strFechaDesde & "' " & _
                " AND P.FECHA_CREACION <= '" & strFechaHasta & "' " & _
                filtro & _
                strTramitado & _
                strTipo & _
                " ORDER BY P.ID_PAQUETE DESC,P.EDICION DESC "
    Set Listado = datos_bd(consulta)
End Function
Public Function recuperarModificado(ID As Long, EDICION As Long) As ADODB.Recordset
    
    Dim consulta As String
    consulta = "SELECT P.CODIGO_SC, PR.NOMBRE CONTRATA, P.PRESUPUESTO  " & _
                ", P.FECHA_CREACION, U.NOMBRE USUARIO , P.ID_PAQUETE, PR.ID_PROVEEDOR, P.ESTADO, P.FECHA_APROBACION, AP.NOMBRE USUARIO, P.TIPO, P.SUBTIPO, temp.suma, P.EDICION, D1.PARAMETROS AS COLOR, D1.DESCRIPCION AS STRESTADO, D2.PARAMETROS, D3.DESCRIPCION AS SUBTIPO,centros.NOMBRE,MON.DESCRIPCION " & _
                " FROM SC_PAQUETES P" & _
                " inner join proveedores PR on p.subcontrata_id = PR.ID_PROVEEDOR" & _
                " inner join usuarios u on P.usuario_id = u.ID_EMPLEADO" & _
                " inner join usuarios AP on P.aprobador_id = AP.ID_EMPLEADO" & _
                " left join decodificadora D1 on P.ESTADO = D1.valor and D1.codigo  = 199" & _
                " left join decodificadora D2 on P.TIPO = D2.valor and D2.codigo  = 500" & _
                " left join decodificadora D3 on P.SUBTIPO = D3.valor and D3.codigo  = 198" & _
                " LEFT JOIN decodificadora MON ON P.MONEDA = MON.VALOR and MON.CODIGO = 180 " & _
                " left join (select tobjeto,cobjeto,proveedor_id, sum(BI) as SUMA from proveedores_facturas WHERE TOBJETO in (31,32,33) GROUP BY tobjeto, cobjeto) as temp on p.id_paquete=temp.cobjeto and p.subcontrata_id = temp.proveedor_id" & _
                " left join centros as centros on p.centro_id = centros.id_centro " & _
                " WHERE P.ID_PAQUETE = " & ID & " AND P.EDICION = " & EDICION
                
    Set recuperarModificado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT ID_PAQUETE, CODIGO_SC FROM SC_PAQUETES ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
        
End Function
Public Function Listado_Script() As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * FROM SC_PAQUETES ORDER BY ID_PAQUETE DESC"
        Set Listado_Script = datos_bd(consulta)
        
End Function
Public Function ListadoProveedor(PROVEEDOR_ID As Long) As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT DISTINCT TIPO,ID_PAQUETE,B.DESCRIPCION,CODIGO_SC,FECHA_CREACION " & _
                   "  FROM SC_PAQUETES A, decodificadora B" & _
                   " WHERE SUBCONTRATA_ID = " & PROVEEDOR_ID & _
                   "   AND A.TIPO = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_TOBJETOS & _
                   "   AND A.ESTADO <> " & SC_ESTADO_HISTORICO & _
                   " ORDER BY ID_PAQUETE DESC"
                   
        Set ListadoProveedor = datos_bd(consulta)
        
End Function
Public Function imprimir(PAQUETE As Long, EDICION As Long, pdf As String) As Boolean
        Dim oPaquete As New clsSC_Paquetes
'M1274-I
'        oPaquete.Carga PAQUETE
        oPaquete.Carga PAQUETE, EDICION
'M1274-F
                
        With frmReport
            .iniciar
            Select Case oPaquete.getTIPO
            Case TOBJETO_SC_DETERMINACIONES
                .informe = "\SC\rptSCPaquetes_Solicitud_Analisis"
            Case TOBJETO_SC_EFICACIA
                .informe = "\SC\rptSCPaquetes_Solicitud_Analisis_CE"
            Case TOBJETO_SC_GENERICA
            Case TOBJETO_SC_PEACH
                .informe = "\SC\rptSCPaquetes_Solicitud_Analisis_GEN"
            End Select
'            .criterio = "{SC_PAQUETES.ID_PAQUETE}=" & PAQUETE & "AND {SC_PAQUETES.EDICION}=" & EDICION
            .criterio = "{sc_paquetes.ID_PAQUETE}=" & PAQUETE & " and {sc_paquetes.EDICION}=" & EDICION & " and {fp.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & " and {vencimiento.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_VENCIMIENTOS
            .imprimir = False
            If Trim(pdf) <> "" Then
              .pdf = pdf
            End If
            .generar
            .visible = False
'            .Show 1
        End With
        Unload frmReport
End Function

Public Function existe_paquete(lngContrata As Long, fecha As String, hora As String) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT ID_PAQUETE " & _
               "  FROM SC_PAQUETES " & _
               " WHERE SUBCONTRATA_ID = " & lngContrata & _
               "   AND FECHA_CREACION = '" & fecha & "'" & _
               "   AND HORA_CREACION = '" & hora & "'" & _
               "   AND TIPO = " & TOBJETO_SC_DETERMINACIONES & _
               "   AND ESTADO <> " & SC_ESTADO_HISTORICO
    Set rs = datos_bd(consulta)
    If Not (IsNull(rs.Fields("ID_PAQUETE")) Or (rs.EOF And rs.BOF)) Then  'si es nulo No se recupero ninguno
        existe_paquete = rs.Fields("ID_PAQUETE")
    Else
        existe_paquete = 0
    End If
    Set rs = Nothing

End Function

'M0957-I
Public Function existe_paquete_CE(lngContrata As Long, fecha As String, hora As String) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT ID_PAQUETE " & _
               "  FROM SC_PAQUETES " & _
               " WHERE SUBCONTRATA_ID = " & lngContrata & _
               "   AND FECHA_CREACION = '" & fecha & "'" & _
               "   AND HORA_CREACION = '" & hora & "'" & _
               "   AND TIPO = " & TOBJETO_SC_EFICACIA & _
               "   AND ESTADO <> " & SC_ESTADO_HISTORICO
                
    Set rs = datos_bd(consulta)
    If Not (IsNull(rs.Fields("ID_PAQUETE")) Or (rs.EOF And rs.BOF)) Then  'si es nulo No se recupero ninguno
        existe_paquete_CE = rs.Fields("ID_PAQUETE")
    Else
        existe_paquete_CE = 0
    End If
    Set rs = Nothing

End Function
'M0957-F

'M1147-I
Public Function existe_paquete_generico(lngContrata As Long, fecha As String, hora As String) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT ID_PAQUETE " & _
               "  FROM SC_PAQUETES " & _
               " WHERE SUBCONTRATA_ID = " & lngContrata & _
               "   AND FECHA_CREACION = '" & fecha & "'" & _
               "   AND HORA_CREACION = '" & hora & "'" & _
               "   AND TIPO = " & TOBJETO_SC_GENERICA & _
               "   AND ESTADO <> " & SC_ESTADO_HISTORICO
                
    Set rs = datos_bd(consulta)
    If Not (IsNull(rs.Fields("ID_PAQUETE")) Or (rs.EOF And rs.BOF)) Then  'si es nulo No se recupero ninguno
        existe_paquete_generico = rs.Fields("ID_PAQUETE")
    Else
        existe_paquete_generico = 0
    End If
    Set rs = Nothing

End Function
'M1147-F
Public Function Listado_muestras_determinaciones(ID_P As String, EDICION As Long) As ADODB.Recordset
    On Error GoTo fallo
    
    Dim consulta As String
    
    consulta = "SELECT CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) CODIGO " & _
               "     , M.REFERENCIA_CLIENTE, TD.NOMBRE DETERMINACION " & _
               "     , PD.VALOR_REFERENCIA, PD.NORMATIVA_APLICABLE,D.ID_DETERMINACION, PD.PRECIO " & _
               "     , M.ID_MUESTRA, TD.ID_TIPO_DETERMINACION ID_TD, D.ID_DETERMINACION " & _
               "  FROM SC_PAQUETES P, SC_PAQUETES_DETALLE PD " & _
               "     , MUESTRAS M, TIPOS_MUESTRA TM " & _
               "     , DETERMINACIONES D, TIPOS_DETERMINACION TD " & _
               " WHERE P.ID_PAQUETE = PD.PAQUETE_ID " & _
               "   AND P.EDICION = PD.EDICION " & _
               "   AND PD.DETERMINACION_ID = D.ID_DETERMINACION " & _
               "   AND M.ID_MUESTRA = D.MUESTRA_ID " & _
               "   AND D.TIPO_DETERMINACION_ID = TD.ID_TIPO_DETERMINACION " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND P.ID_PAQUETE = " & ID_P & _
               "   AND P.EDICION = " & EDICION & _
               " ORDER BY CODIGO "
    
    Set Listado_muestras_determinaciones = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function

'M0957-I
Public Function Listado_muestras_ensayos(ID_P As String, EDICION As Long) As ADODB.Recordset
    On Error GoTo fallo
    
    Dim consulta As String
    
    consulta = "SELECT CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) CODIGO " & _
               "     , M.REFERENCIA_CLIENTE, TE.NOMBRE " & _
               "     , PD.VALOR_REFERENCIA, PD.NORMATIVA_APLICABLE, M.ID_MUESTRA, pd.TIPO_ENSAYO_ID ID_TE, PD.PRECIO " & _
               "  FROM SC_PAQUETES P, SC_PAQUETES_DETALLE PD " & _
               "     , MUESTRAS M, TIPOS_MUESTRA TM " & _
               "     , CE_TIPOS_ENSAYOS TE " & _
               " WHERE P.ID_PAQUETE = PD.PAQUETE_ID " & _
               "   AND P.EDICION = PD.EDICION " & _
               "   AND PD.MUESTRA_ID = M.ID_MUESTRA " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND PD.TIPO_ENSAYO_ID = TE.ID_TIPO_ENSAYO " & _
               "   AND P.ID_PAQUETE = " & ID_P & _
               "   AND P.EDICION = " & EDICION & _
               " ORDER BY CODIGO "
    
    Set Listado_muestras_ensayos = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function
'M0957-F

Public Function paqueteAsociadoMuestra(idMuestra As Long) As String
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo paqueteAsociadoMuestra_Error

    consulta = "select DISTINCT concat(a.PAQUETE_ID,'/',a.EDICION) from sc_paquetes_detalle a,determinaciones b " & _
                " where a.DETERMINACION_ID = b.ID_DETERMINACION " & _
                " and b.MUESTRA_ID = " & idMuestra & _
                " Union  " & _
                " select DISTINCT concat(a.PAQUETE_ID,'/',a.EDICION) from sc_paquetes_detalle a " & _
                " where a.MUESTRA_ID = " & idMuestra & _
                " order by 1 desc "
    Set rs = datos_bd(consulta)
    If Not (IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF)) Then  'si es nulo No se recupero ninguno
        paqueteAsociadoMuestra = rs(0)
    Else
        paqueteAsociadoMuestra = ""
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Function

paqueteAsociadoMuestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure paqueteAsociadoMuestra of Módulo de clase clsSC_Paquetes"

End Function


