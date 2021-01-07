VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private ID_DOCUMENTO As Long
Private TIPO_DOCUMENTO_ID As Long
Private numero As Long
Private anno As Integer
Private OBRA_ID As Long
Private fecha As String
Private FECHA_COBRO As String
Private total As String
Private descuento As String
Private DESCUENTO_PORCENTAJE As String
Private portes As String
Private facturado As Integer
Private DOCUMENTO_ID_REL As Long
Private OBSERVACIONES  As String
Private TARIFA_ID As Integer
Private FP_ID As Integer
Private SERVIDO As String
Private VALORACION As Integer
Private VEHICULO_ID As Integer
Private MATRICULA As String
Private Nif As String
Private REMOLQUE As String
Private USUARIO_ID As Long
Private hora As String
Private BULTOS As Long
Private PESO As Long
Private iva As Integer
Private anulado As Integer
Private CONTABILIZADO As Integer
Private CONTABILIZADO_BM As String
Private ESTADO_ID As Integer
Public Property Let setID_DOCUMENTO(ByVal dato As Long)
        ID_DOCUMENTO = dato
End Property
Public Property Let setTIPO_DOCUMENTO_ID(ByVal dato As Long)
        TIPO_DOCUMENTO_ID = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
        numero = dato
End Property
Public Property Let setANNO(ByVal dato As Integer)
        anno = dato
End Property
Public Property Let setOBRA_ID(ByVal dato As Long)
        OBRA_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setFECHA_COBRO(ByVal dato As String)
        FECHA_COBRO = dato
End Property
Public Property Let setTOTAL(ByVal dato As String)
        total = dato
End Property
Public Property Let setDESCUENTO(ByVal dato As String)
        descuento = dato
End Property
Public Property Let setDESCUENTO_PORCENTAJE(ByVal dato As String)
        DESCUENTO_PORCENTAJE = dato
End Property
Public Property Let setPORTES(ByVal dato As String)
        portes = dato
End Property
Public Property Let setFACTURADO(ByVal dato As Integer)
        facturado = dato
End Property
Public Property Let setDOCUMENTO_ID_REL(ByVal dato As Long)
        DOCUMENTO_ID_REL = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setTARIFA_ID(ByVal dato As Integer)
        TARIFA_ID = dato
End Property
Public Property Let setFP_ID(ByVal dato As Integer)
        FP_ID = dato
End Property
Public Property Let setSERVIDO(ByVal dato As String)
        SERVIDO = dato
End Property
Public Property Let setVALORACION(ByVal dato As Integer)
        VALORACION = dato
End Property
Public Property Let setVEHICULO_ID(ByVal dato As Integer)
        VEHICULO_ID = dato
End Property
Public Property Let setMATRICULA(ByVal dato As String)
        MATRICULA = dato
End Property
Public Property Let setNIF(ByVal dato As String)
        Nif = dato
End Property
Public Property Let setREMOLQUE(ByVal dato As String)
        REMOLQUE = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
Public Property Let setBULTOS(ByVal dato As Long)
        BULTOS = dato
End Property
Public Property Let setPESO(ByVal dato As Long)
        PESO = dato
End Property
Public Property Let setIVA(ByVal dato As Integer)
        iva = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
        anulado = dato
End Property
Public Property Let setCONTABILIZADO(ByVal dato As Integer)
        CONTABILIZADO = dato
End Property
Public Property Let setCONTABILIZADO_BM(ByVal dato As String)
        CONTABILIZADO_BM = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
        ESTADO_ID = dato
End Property
'*************************************
' GETTER
'*************************************
Public Property Get getID_DOCUMENTO() As Long
        getID_DOCUMENTO = ID_DOCUMENTO
End Property
Public Property Get getTIPO_DOCUMENTO_ID() As Long
        getTIPO_DOCUMENTO_ID = TIPO_DOCUMENTO_ID
End Property
Public Property Get getNUMERO() As Long
        getNUMERO = numero
End Property
Public Property Get getANNO() As Integer
        getANNO = anno
End Property
Public Property Get getOBRA_ID() As Long
        getOBRA_ID = OBRA_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getFECHA_COBRO() As String
        getFECHA_COBRO = FECHA_COBRO
End Property
Public Property Get getTOTAL() As String
        getTOTAL = total
End Property
Public Property Get getDESCUENTO() As String
        getDESCUENTO = descuento
End Property
Public Property Get getDESCUENTO_PORCENTAJE() As String
        getDESCUENTO_PORCENTAJE = DESCUENTO_PORCENTAJE
End Property
Public Property Get getPORTES() As String
        getPORTES = portes
End Property
Public Property Get getFACTURADO() As Integer
        getFACTURADO = facturado
End Property
Public Property Get getDOCUMENTO_ID_REL() As Long
        getDOCUMENTO_ID_REL = DOCUMENTO_ID_REL
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getTARIFA_ID() As Integer
        getTARIFA_ID = TARIFA_ID
End Property
Public Property Get getFP_ID() As Integer
        getFP_ID = FP_ID
End Property
Public Property Get getSERVIDO() As String
        getSERVIDO = SERVIDO
End Property
Public Property Get getVALORACION() As Integer
        getVALORACION = VALORACION
End Property
Public Property Get getVEHICULO_ID() As Integer
        getVEHICULO_ID = VEHICULO_ID
End Property
Public Property Get getMATRICULA() As String
        getMATRICULA = MATRICULA
End Property
Public Property Get getNIF() As String
        getNIF = Nif
End Property
Public Property Get getREMOLQUE() As String
        getREMOLQUE = REMOLQUE
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
Public Property Get getBULTOS() As Long
        getBULTOS = BULTOS
End Property
Public Property Get getPESO() As Long
        getPESO = PESO
End Property
Public Property Get getIVA() As Integer
        getIVA = iva
End Property
Public Property Get getANULADO() As Integer
        getANULADO = anulado
End Property
Public Property Get getCONTABILIZADO() As Integer
        getCONTABILIZADO = CONTABILIZADO
End Property
Public Property Get getCONTABILIZADO_BM() As String
        getCONTABILIZADO_BM = CONTABILIZADO_BM
End Property
Public Property Get getESTADO_ID() As Integer
        getESTADO_ID = ESTADO_ID
End Property
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_DOCUMENTO) FROM DOCUMENTOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_DOCUMENTO = 1
    Else
        ID_DOCUMENTO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Sub Calcular_Numero_factura()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(NUMERO) FROM DOCUMENTOS WHERE ANNO=" & anno & " AND TIPO_DOCUMENTO_ID=" & TIPO_DOCUMENTO_ID
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        numero = 1
    Else
        numero = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS WHERE ID_DOCUMENTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_DOCUMENTO = rs("ID_DOCUMENTO")
                TIPO_DOCUMENTO_ID = rs("TIPO_DOCUMENTO_ID")
                anno = rs("ANNO")
                numero = rs("NUMERO")
                OBRA_ID = rs("OBRA_ID")
                fecha = rs("FECHA")
                If Not IsNull(rs("FECHA_COBRO")) Then
                    FECHA_COBRO = rs("FECHA_COBRO")
                End If
                total = rs("TOTAL")
                descuento = rs("DESCUENTO")
                DESCUENTO_PORCENTAJE = rs("DESCUENTO_PORCENTAJE")
                portes = rs("PORTES")
                facturado = rs("FACTURADO")
                DOCUMENTO_ID_REL = rs("DOCUMENTO_ID_REL")
                OBSERVACIONES = rs("OBSERVACIONES")
                TARIFA_ID = rs("TARIFA_ID")
                FP_ID = rs("FP_ID")
                SERVIDO = rs("SERVIDO")
                VALORACION = rs("VALORACION")
                VEHICULO_ID = rs("VEHICULO_ID")
                MATRICULA = rs("MATRICULA")
                Nif = rs("NIF")
                REMOLQUE = rs("REMOLQUE")
                USUARIO_ID = rs("USUARIO_ID")
                hora = rs("HORA")
                BULTOS = rs("BULTOS")
                PESO = rs("PESO")
                iva = rs("IVA")
                anulado = rs("ANULADO")
                CONTABILIZADO = rs("CONTABILIZADO")
                If Not IsNull(rs("CONTABILIZADO_BM")) Then
                    CONTABILIZADO_BM = rs("CONTABILIZADO_BM")
                End If
                ESTADO_ID = rs("ESTADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function Carga_Ultimo_Albaran(OBRA As Long, ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        If ID <> 0 Then
            s = s & " AND ID_DOCUMENTO < " & ID
        End If
        consulta = "SELECT * FROM DOCUMENTOS WHERE TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN & " AND OBRA_ID = " & OBRA & s & " ORDER BY ID_DOCUMENTO DESC"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_Ultimo_Albaran = False
        Else
                ID_DOCUMENTO = rs("ID_DOCUMENTO")
                anno = rs("ANNO")
                numero = rs("NUMERO")
                fecha = rs("FECHA")
                Carga_Ultimo_Albaran = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_Ultimo_Albaran = False
        MsgBox "Error al Carga_Ultimo_Albaran los datos(clsDOCUMENTOS)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_DOCUMENTO = 0 Then
            Me.CrearID
        End If
        If numero = 0 Then
            Me.Calcular_Numero_factura
        End If
        Dim oObra As New clsObras
        Dim ocliente As New clsCliente
        
        If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
            Me.setIVA = 0
        Else
            Me.setIVA = ReadINI(App.Path & "\config.ini", "parametros", "iva")
            If oObra.Carga(CLng(OBRA_ID)) Then
                If oObra.getTIPO_IVA = 5 Then
                    Me.setIVA = 0
                End If
    '            If ocliente.CargaCliente(CLng(oObra.getCLIENTE_ID)) = True Then
    '                Me.setIVA = ocliente.getIVA
    '            End If
            End If
        End If
        Me.setDESCUENTO = moneda_bd("0")
        Me.setDESCUENTO_PORCENTAJE = moneda_bd("0")
        FECHA_COBRO = "0000-00-00"
' Modificar las comillas para el numero en alfanumerico
        consulta = "INSERT INTO DOCUMENTOS " & _
                   "  VALUES (" & _
                        ID_DOCUMENTO & "," & TIPO_DOCUMENTO_ID & "," & anno & "," & numero & "," & _
                        OBRA_ID & "," & _
                        "'" & fecha & "','" & FECHA_COBRO & "','" & total & "','" & descuento & "','" & DESCUENTO_PORCENTAJE & "','" & portes & "'," & _
                        facturado & "," & DOCUMENTO_ID_REL & ",'" & _
                        OBSERVACIONES & "'," & _
                        TARIFA_ID & "," & FP_ID & ",'" & SERVIDO & "'," & VALORACION & "," & VEHICULO_ID & ",'" & _
                        MATRICULA & "','" & Nif & "','" & REMOLQUE & "'," & USUARIO_ID & ",'" & hora & "'," & _
                        BULTOS & "," & PESO & "," & iva & "," & anulado & "," & CONTABILIZADO & ",'0000-00-00 00:00:00'," & ESTADO_ID & ")"
        execute_bd consulta
        Me.Calcular_Descuento ID_DOCUMENTO
        Insertar = ID_DOCUMENTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " OBRA_ID = " & OBRA_ID & "," & _
                        " TARIFA_ID = " & TARIFA_ID & "," & _
                        " FP_ID = " & FP_ID & "," & _
                        " SERVIDO = '" & SERVIDO & "'," & _
                        " VALORACION = " & VALORACION & "," & _
                        " VEHICULO_ID = " & VEHICULO_ID & "," & _
                        " BULTOS = " & BULTOS & "," & _
                        " PESO = " & PESO & "," & _
                        " MATRICULA = '" & MATRICULA & "'," & _
                        " NIF = '" & Nif & "'," & _
                        " REMOLQUE = '" & REMOLQUE & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " TOTAL = '" & total & "'," & _
                        " PORTES = '" & portes & "'" & _
                " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Me.Calcular_Descuento (ID)
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function Calcular_Descuento(ID As Long) As Boolean
        Dim oDOCUMENTO As New clsDocumentos
        If oDOCUMENTO.Carga(ID) = True Then
'            If oDOCUMENTO.getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.factura Then
                If oDOCUMENTO.getTOTAL > 0 Then
                    Dim oObra As New clsObras
                    If oObra.Carga(oDOCUMENTO.getOBRA_ID) = True Then
                        If CSng(oObra.getDESCUENTO) <> 0 Then
                            Dim cDescuento As Currency
                            cDescuento = (oDOCUMENTO.getTOTAL * oObra.getDESCUENTO) / 100
                            execute_bd "UPDATE DOCUMENTOS " & _
                                       "   SET DESCUENTO = '" & Replace(Str(cDescuento), ",", ".") & "'" & _
                                       "      ,DESCUENTO_PORCENTAJE = '" & Replace(oObra.getDESCUENTO, ",", ".") & "'" & _
                                       " WHERE ID_DOCUMENTO = " & ID
                        End If
                    End If
                    Set oObra = Nothing
                End If
'            End If
        End If
        Set oDOCUMENTO = Nothing
End Function

Public Function Anular(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        ' Verificar si ya esta anulado
        consulta = "SELECT ANULADO FROM DOCUMENTOS WHERE ID_DOCUMENTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            If rs(0) = 1 Then
                MsgBox "El documento ya se encuentra anulado.", vbCritical, App.Title
                Anular = False
                Exit Function
            End If
        End If
        ' Anular
        execute_bd "UPDATE DOCUMENTOS SET FACTURADO = 0 WHERE DOCUMENTO_ID_REL = " & ID
        execute_bd "UPDATE DOCUMENTOS SET ANULADO = 1, ESTADO_ID = " & ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_ANULADO & " WHERE ID_DOCUMENTO = " & ID
        Anular = True
        Exit Function
fallo:
        Anular = False
        MsgBox "Error al Anular (clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function Contabilizar(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "UPDATE DOCUMENTOS SET CONTABILIZADO = 1 WHERE ID_DOCUMENTO = " & ID
        Contabilizar = True
        Exit Function
fallo:
        Contabilizar = False
        MsgBox "Error al Contabilizar (clsDOCUMENTOS)", vbCritical, Err.Description
End Function

Public Function Contabilizar_BM(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "UPDATE DOCUMENTOS SET CONTABILIZADO_BM = CURRENT_TIMESTAMP WHERE ID_DOCUMENTO = " & ID
        Contabilizar_BM = True
        Exit Function
fallo:
        Contabilizar_BM = False
        MsgBox "Error al Contabilizar (clsDOCUMENTOS)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS_DETALLE " & _
                   " WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM DOCUMENTOS " & _
                   " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS ORDER BY ID_DOCUMENTO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_cliente(cliente As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT D.FECHA,TD.NOMBRE,D.NUMERO,C.NOMBRE,D.TOTAL,D.ID_DOCUMENTO " & _
                   "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C " & _
                   " WHERE D.OBRA_ID = O.ID_OBRA " & _
                   "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
                   "   AND O.CLIENTE_ID = " & cliente & _
                   " ORDER BY D.FECHA DESC"
        Set Listado_por_cliente = datos_bd(consulta)
End Function
Public Function Listado_albaranes_por_factura(factura As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT D.FECHA,TD.NOMBRE,D.NUMERO,C.NOMBRE,D.TOTAL,D.ID_DOCUMENTO " & _
                   "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C " & _
                   " WHERE D.OBRA_ID = O.ID_OBRA " & _
                   "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
                   "   AND D.DOCUMENTO_ID_REL = " & factura & _
                   "   AND D.TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN & _
                   " ORDER BY D.FECHA DESC"
        Set Listado_albaranes_por_factura = datos_bd(consulta)
End Function
Public Function Listado_por_obra(OBRA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT D.FECHA,TD.NOMBRE,D.NUMERO,C.NOMBRE,D.TOTAL,D.ID_DOCUMENTO " & _
                   "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C " & _
                   " WHERE D.OBRA_ID = O.ID_OBRA " & _
                   "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
                   "   AND O.ID_OBRA = " & OBRA & _
                   " ORDER BY D.FECHA DESC"
        Set Listado_por_obra = datos_bd(consulta)
End Function
'Public Function Comprobar_duplicado(numero As Double, ANNO As Integer, tipo As Integer) As Boolean
Public Function Comprobar_duplicado(numero As String, anno As Integer, tipo As Integer) As Boolean
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT ID_DOCUMENTO " & _
                   "  FROM DOCUMENTOS " & _
                   " WHERE NUMERO = " & numero & _
                   "   AND ANNO = " & anno & _
                   "   AND TIPO_DOCUMENTO_ID = " & tipo
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Comprobar_duplicado = False
        Else
            Comprobar_duplicado = True
        End If
End Function

Public Function modificar_numero(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS SET " & _
                        " ANNO = " & anno & "," & _
                        " NUMERO = " & numero & _
                " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        modificar_numero = True
        Exit Function
fallo:
        modificar_numero = False
        MsgBox "Error al Modificar el número de documento(clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function COBRADO(ID As Long, fecha As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS " & _
                   "   SET FECHA_COBRO = '" & Format(fecha, "YYYY-MM-DD") & "'" & _
                   " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        COBRADO = True
        Exit Function
fallo:
        COBRADO = False
        MsgBox "Error EN EL cobrado deL documento(clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function modificar_estado(ID As Long, ESTADO As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS " & _
                   "   SET ESTADO_ID = " & ESTADO & _
                   " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        modificar_estado = True
        Exit Function
fallo:
        modificar_estado = False
        MsgBox "Error al modificar_estado deL documento(clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function modificar_fecha(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS SET " & _
                   " FECHA = '" & fecha & "' " & _
                   " WHERE ID_DOCUMENTO = " & ID
        execute_bd consulta
        modificar_fecha = True
        Exit Function
fallo:
        modificar_fecha = False
        MsgBox "Error al Modificar la fecha del documento(clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function facturar(OBRA As Long, grupo As String, DOC As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS " & _
                   "   SET FACTURADO = 1, DOCUMENTO_ID_REL = " & DOC & _
                   " WHERE OBRA_ID = " & OBRA & _
                   "   AND FACTURADO = 0 " & _
                   "   AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN & _
                   "   AND ID_DOCUMENTO IN (" & grupo & ")"
        execute_bd consulta
        facturar = True
        Exit Function
fallo:
        facturar = False
        MsgBox "Error al pasar a facturado el documento(clsDOCUMENTOS)", vbCritical, Err.Description
End Function
Public Function Listado_documentos_para_factura(grupo As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT OBRA_ID, SUM(TOTAL),SUM(PORTES) " & _
                   "  FROM DOCUMENTOS " & _
                   " WHERE ID_DOCUMENTO IN (" & grupo & ")" & _
                   "   AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN & _
                   " GROUP BY OBRA_ID " & _
                   " ORDER BY OBRA_ID, ID_DOCUMENTO"
        Set Listado_documentos_para_factura = datos_bd(consulta)
End Function
Public Function Listado_detalle_documentos_para_factura(OBRA, grupo As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "select dd.articulo_id,d.fecha,d.numero,dd.cantidad,dd.descripcion,dd.precio,dd.total,dd.portes,d.servido " & _
                   "  from documentos d, documentos_detalle dd " & _
                   " where d.id_documento = dd.documento_id " & _
                   "   and d.obra_id = " & OBRA & _
                   "   and d.tipo_documento_id = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN & _
                   "   and d.facturado = 0 " & _
                   "   and d.id_documento in (" & grupo & ")" & _
                   " order by d.id_documento, dd.orden"
        Set Listado_detalle_documentos_para_factura = datos_bd(consulta)
End Function

'Public Function modificar_total(ID As Long) As Boolean
'        On Error GoTo fallo
'        Dim consulta As String
'        consulta = "UPDATE DOCUMENTOS SET " & _
'                   " TOTAL = '" & total & "'" & _
'                   " WHERE ID_DOCUMENTO = " & ID
'        execute_bd consulta
'        modificar_total = True
'        Exit Function
'fallo:
'        modificar_total = False
'        MsgBox "Error al Modificar el total (clsDOCUMENTOS)", vbCritical, Err.Description
'End Function
'Public Sub Recalcular_TOTAL(Documento As Long)
'    Dim DTO As Currency
'    Dim oDocumento As New clsDocumentos
'    Dim oDocumento_Detalle As New clsDocumentos_detalle
'    Dim rs As ADODB.Recordset
'   On Error GoTo calcular_TOTAL_Error
'    Dim total As Single
'    Dim IVA As Single
'    IVA = 0
'    total = 0
'    If oDocumento.Carga(Documento) = True Then
'        ' Base e iva
'        Set rs = oDocumento_Detalle.Detalle_Documento(Documento)
'        If rs.RecordCount > 0 Then
'            Do
'                total = total + CSng(rs(4))
'                IVA = IVA + (((CSng(rs(4))) * CSng(rs(7))) / 100)
'                rs.MoveNext
'            Loop Until rs.EOF
'        End If
'        ' Cargo
'        Dim CARGO As Single
'        CARGO = 0
''        CARGO = CSng(oDOCUMENTO.getCARGO)
'        ' DTO
''        If oDOCUMENTO.getDTO1 > 0 Then
''            DTO = Format((CCur(total) * oDOCUMENTO.getDTO1 / 100), "#,##0.00")
''        Else
'            DTO = 0
''        End If
''        DTO = DTO + Format(((CCur(total) - DTO) * oDOCUMENTO.getDTO2 / 100), "#,##0.00")
'        ' Total
'        Dim totalt As String
'        totalt = Format(CCur(total) + CCur(CARGO) - CCur(DTO) + CCur(IVA), "currency")
'        totalt = Replace(Format(totalt, "0.00"), ",", ".")
'        oDocumento.setTOTAL = totalt
'        oDocumento.modificar_total (Documento)
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'calcular_TOTAL_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcular_TOTAL of Formulario frmimprimir"
'End Sub
Public Sub imprimir(Documento As Long, impresora As Boolean, logo As Boolean, copias As Integer, Optional pdf As String = "", Optional ESCOPIA As Boolean)
    Dim oDOCUMENTO As New clsDocumentos
'    Dim tipo As String
   On Error GoTo Imprimir_Error

    oDOCUMENTO.Carga Documento
    Dim aux As String
    Dim p1() As String
    Dim p2() As String
    With frmReport
        .iniciar
        If oDOCUMENTO.getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN Then
            .informe = "rptAlbaran"
            aux = " AND {decodificadora.CODIGO} = " & DECODIFICADORA.D_TIPOS_FACTURACION
            ReDim p1(2) As String
            ReDim p2(2) As String
            p1(1) = "LOGO"
            If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
'            If logo Then
                p2(1) = "N"
            Else
                p2(1) = "S"
            End If
            p1(2) = "COPIA"
            p2(2) = "N"
        Else
            ReDim p1(1) As String
            ReDim p2(1) As String
            p1(1) = "COPIA"
            If ESCOPIA Then
                p2(1) = "S"
            Else
                p2(1) = "N"
            End If
                
            aux = " AND {documentos_detalle.TOTAL} <> 0.00 "
            
            If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                .informe = "rptFacturaPequena"
            Else
                If logo Then
                    .informe = "rptFacturaCompleta"
                Else
                    .informe = "rptFactura"
                End If
            End If
        End If
        .ParametrosNombre = p1
        .ParametrosValores = p2
        .CRITERIO = "{documentos.ID_DOCUMENTO}=" & Documento & aux
        If pdf <> "" Then
            .pdf = pdf
        End If
        .imprimir = impresora
        .copias = copias
         If .generar = True Then
            If pdf = "" And impresora = False Then
                .Show 1
            End If
        End If
    End With
    Unload frmReport
    
    'Si el albaran y sin logo, preguntamos si se quiere copia oficial
    If oDOCUMENTO.getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN And frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
        If MsgBox("¿Desea imprimir la copia?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                With frmReport
                    .iniciar
                    .informe = "rptAlbaran"
                    aux = " AND {decodificadora.CODIGO} = " & DECODIFICADORA.D_TIPOS_FACTURACION
                    ReDim p1(2) As String
                    ReDim p2(2) As String
                    p1(1) = "LOGO"
                    p2(1) = "S"
                    p1(2) = "COPIA"
                    p2(2) = "S"
                    .ParametrosNombre = p1
                    .ParametrosValores = p2
                    .CRITERIO = "{DOCUMENTOS.ID_DOCUMENTO}=" & Documento & aux
                    If pdf <> "" Then
                        .pdf = pdf
                    End If
                    .imprimir = impresora
                    .copias = copias
                    .generar

                    If pdf = "" And impresora = False Then
                        .Show 1
                    End If
            End With
            Unload frmReport
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

Imprimir_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Imprimir of Módulo de clase clsDocumentos"
End Sub

Public Function Correo(Documento As Long, impresora As Boolean, logo As Boolean, copias As Integer, Optional ESCOPIA As Boolean) As Boolean
    
    Dim oDOC As New clsDocumentos
   On Error GoTo correo_Error

    With oDOC
    If .Carga(Documento) Then
        Dim NOMBRE As String
        If .getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN Then
            NOMBRE = "Albaran " & .getNUMERO & "-" & .getANNO & ".pdf"
        Else
            NOMBRE = "Factura " & .getNUMERO & "-" & .getANNO & ".pdf"
        End If
        Dim destino_documento As String
        On Error Resume Next
        MkDir App.Path & "\tmp\"
        destino_documento = App.Path & "\tmp\" & NOMBRE
        If Dir(destino_documento) <> "" Then
            Kill destino_documento
        End If
   On Error GoTo correo_Error
        ' Generamos el pdf
        Me.imprimir Documento, impresora, logo, copias, destino_documento, ESCOPIA
        ' Obtenemos los datos del correo
        Dim oObra As New clsObras
        oObra.Carga oDOC.getOBRA_ID
        Dim ocliente As New clsCliente
        ocliente.CargaCliente oObra.getCLIENTE_ID
        
        Dim ref As String
        If Dir(destino_documento) = "" Then
            MsgBox "El informe con la factura no se ha generado correctamente.", vbInformation, App.Title
            Correo = False
            Exit Function
        End If
        
        If .getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN Then
            ref = "Adjunto albarán número: " & oDOC.getNUMERO & "-" & oDOC.getANNO
        Else
            ref = "Adjunto factura número: " & oDOC.getNUMERO & "-" & oDOC.getANNO
        End If
        genera_correo_outlook ocliente.getEMAIL, ref, "", destino_documento, 0
    End If
    End With
    Set oDOC = Nothing
    Correo = True

   On Error GoTo 0
   Exit Function

correo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure correo of Módulo de clase clsDocumentos"
End Function

Public Function actualizar_Albaran_en_Factura(idAlbaran As Long) As Boolean

On Error GoTo errorActualizacion

    Dim odetalleDocumento As New clsDocumentos_detalle
    Dim oAlbaran As New clsDocumentos
    Dim nuevoTotal As Long
    
    log "actualizar_Albaran_en_Factura: Actualizando el detalle del albaran " & idAlbaran & " en la factura " & getID_DOCUMENTO
    
    actualizar_Albaran_en_Factura = False
    
    'Carga los datos del albaran que ha sido modificado
    If Not oAlbaran.Carga(idAlbaran) Then
        MsgBox "Error en la carga del albaran para su modificación en la factura", vbExclamation, App.Title
        Exit Function
    End If
    
    
    'Elimina el detalle del albaran modificado de la factura
    log "actualizar_Albaran_en_Factura: Borrado de las lineas de detalle del albaran"
    
    borrar_Detalle_Albaran getID_DOCUMENTO, oAlbaran.getNUMERO
    
    'Inserta al final las lineas de detalle del albarán que se ha modificado.
    log "actualizar_Albaran_en_Factura: Inserta el nuevo detalle"
    inserta_Detalle_Albaran getID_DOCUMENTO, oAlbaran.getID_DOCUMENTO
    
    'Carga el detalle actual de la factura
    If Not odetalleDocumento.Carga(getID_DOCUMENTO) Then
        MsgBox "Error en la carga del detalle de la factura para su modificación", vbExclamation, App.Title
        Exit Function
    End If
    
    log "actualizar_Albaran_en_Factura: Reordenacion de las lineas de detalle"
    odetalleDocumento.reordenar_Detalle getID_DOCUMENTO
    
    'se recalcula el importe y el descuento de la factura en base al nuevo importe por la modificación del detalle.
    log "actualizar_Albaran_en_Factura: Recalcular el importe final"
    nuevoTotal = recalcular_Factura(getID_DOCUMENTO)
    
    If nuevoTotal <> getTOTAL Then
        MsgBox "El total de la factura ha cambiado, no olvide revisar los recibos de la factura", vbInformation, App.Title
    End If
    
    actualizar_Albaran_en_Factura = True
Exit Function
errorActualizacion:

actualizar_Albaran_en_Factura = False

End Function
Private Function borrar_Detalle_Albaran(idFactura As Long, numeroAlbaran As Long) As Boolean

    Dim consulta As String
    
    On Error GoTo ErrorEjecucion
    
    borrar_Detalle_Albaran = False
    
    consulta = "DELETE FROM DOCUMENTOS_DETALLE " & _
                "WHERE DOCUMENTO_ID = " & idFactura & _
                " AND NUMERO_ALBARAN = " & numeroAlbaran
                
    bd.execute_bd consulta
    borrar_Detalle_Albaran = True
    Exit Function
ErrorEjecucion:
    MsgBox "Error en la actualización de la factura " & Chr(13) & _
    " borrar_Detalle_Albaran: " & Err.Description, vbExclamation, App.Title
End Function

Private Function inserta_Detalle_Albaran(idFactura As Long, idAlbaran As Long) As Boolean
    On Error GoTo ErrorEjecucion
    
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim maximo As Single
    Dim oAlbaran As New clsDocumentos
    
    inserta_Detalle_Albaran = False
    
    consulta = "SELECT 1 FROM DOCUMENTOS_DETALLE " & _
                "WHERE DOCUMENTO_ID = " & idFactura
                
    Set rs = bd.datos_bd(consulta)
    
    If rs.RecordCount = 0 Then
        maximo = 0
    Else
        Set rs = Nothing
        
        consulta = "SELECT MAX(ORDEN) FROM DOCUMENTOS_DETALLE " & _
                    "WHERE DOCUMENTO_ID = " & idFactura
                    
        Set rs = bd.datos_bd(consulta)
                
        maximo = rs(0) + rs.RecordCount + 1
    End If
    
    Set rs = Nothing
    
    oAlbaran.Carga idAlbaran

    consulta = "INSERT INTO DOCUMENTOS_DETALLE " & _
                "(SELECT " & getID_DOCUMENTO & _
                ", ORDEN + " & maximo & _
                ", ARTICULO_ID,'" & Format(oAlbaran.getFECHA, "dd-mm-yyyy") & _
                "'," & oAlbaran.getNUMERO & ",'" & oAlbaran.getSERVIDO & _
                "', CANTIDAD, DESCRIPCION, PRECIO, TOTAL, PORTES FROM DOCUMENTOS_DETALLE WHERE DOCUMENTO_ID = " & idAlbaran & ")"
                
    bd.execute_bd consulta
    
    inserta_Detalle_Albaran = True
    Exit Function
ErrorEjecucion:
    MsgBox "Error en la actualización de la factura " & Chr(13) & _
    " inserta_Detalle_Albaran: " & Err.Description, vbExclamation, App.Title
End Function

Private Function recalcular_Factura(idFactura As Long) As Long

    On Error GoTo ErrorEjecucion
    
    Dim portes As String
    Dim total As String
    Dim consulta As String
    Dim rs As ADODB.Recordset
        
    consulta = "SELECT SUM(TOTAL) AS TOTAL_FACTURA,SUM(PORTES) AS TOTAL_PORTES " & _
                "FROM DOCUMENTOS_DETALLE " & _
                "WHERE DOCUMENTO_ID = " & idFactura
    
    Set rs = bd.datos_bd(consulta)
    
    If rs.RecordCount > 0 Then
        total = rs("TOTAL_FACTURA")
        portes = rs("TOTAL_PORTES")
    Else
        total = 0
        portes = 0
    End If
    
    Set rs = Nothing
    
    consulta = "UPDATE DOCUMENTOS " & _
                "SET TOTAL = " & moneda_bd(total) & _
                ", PORTES = " & moneda_bd(portes) & _
                ", DESCUENTO = ((TOTAL * descuento_porcentaje) / 100) " & _
                " WHERE ID_DOCUMENTO = " & idFactura
    
    bd.execute_bd consulta
    
    recalcular_Factura = total
    
    Exit Function
ErrorEjecucion:
    MsgBox "Error en la actualización de la factura " & Chr(13) & _
    " recalcularFactura: " & Err.Description, vbExclamation, App.Title
End Function
