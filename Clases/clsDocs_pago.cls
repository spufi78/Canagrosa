VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocs_pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_DOC As Long
Private TIPO As Integer
Private NUMERO As Long
Private FECHA_FACTURA As String
Private FECHA_GENERACION As String
Private EMPLEADO_ID As Integer
Private CLIENTE_ID As Long
Private CLIENTE_ID_FACTURA As Long
Private total As String
Private DESCUENTO As String
Private IVA As Integer
Private pagado As Long
Private factura_conceptos As Integer
Private FP_ID As Integer
Private PEDIDO_ID As Integer
Private CONTABILIZADA As String
Private FACTORIZADA As String
Private OBSERVACIONES As String
Private COMENTARIO As String
Private anulado As Integer
Private ENVIADA As Integer
Private asiento As Long
Public Function getNUMERO_FORMATEADO() As String
    Dim NUMERO As String
    Select Case Me.getTIPO
    Case 1
        NUMERO = "A-" & Format(Me.getNUMERO, "0000")
    Case 2
        NUMERO = "F-" & Format(Me.getNUMERO, "0000")
    Case 3
        NUMERO = "B-" & Format(Me.getNUMERO, "0000")
    Case 4
        NUMERO = "P-" & Format(Me.getNUMERO, "0000")
    Case Else
        NUMERO = Format(Me.getNUMERO, "0000")
    End Select
    getNUMERO_FORMATEADO = NUMERO
End Function

Public Function ImprimirListado(cliente As Long, clienteFact As Long, fdesde As String, fhasta As String, lTIPO As String, lpagado As Integer, ESTADO As String, ENVIADA As String, AIRBUS As Boolean, vencimiento As Boolean, fdesdev As String, fhastav As String, IBERIA As Boolean) As ADODB.Recordset

    Dim objfrm As New frmReport
    Dim strCad As String
    Dim arrNom() As String
    Dim arrVal() As String
    
   On Error GoTo ImprimirListado_Error

    With objfrm
        .iniciar
        .informe = "Facturacion\rptListadoDocPago_ClientesTodos"
'        .informe = "Facturacion\rptListadoDocPago"
        ' Cliente
        strCad = " 1 = 1 "
        If cliente <> 0 Then
            strCad = strCad & " AND {docs_pago.CLIENTE_ID} =  " & cliente
        End If
        ' ClienteFact
        If clienteFact <> 0 Then
            strCad = strCad & " AND {docs_pago.CLIENTE_ID_FACTURA} =  " & clienteFact
        End If
        ' ltipo
'        If lTIPO <> 0 Then
'            strCad = strCad & " and {docs_pago.TIPO} = " & CStr(lTIPO)
        If lTIPO <> "" Then
            strCad = strCad & " and {docs_pago.TIPO} in [" & CStr(lTIPO) & "] "
        Else
            If lpagado = 0 Then
                strCad = strCad & " and {docs_pago.TIPO} = 2 "
            Else
                strCad = strCad & " and {docs_pago.TIPO} in [2,3] "
            End If
        End If
        'lpagado
        If lpagado < 2 Then
            strCad = strCad & " and {docs_pago.PAGADO} = " & CStr(lpagado)
        End If
        ' ESTADO
        strCad = strCad & " and {docs_pago.ANULADO} in [" & ESTADO & "]"
        ' AIRBUS
        If AIRBUS = True Then
            strCad = strCad & " AND {clientes.AIRBUS} = 1 "
        End If
        ' IBERIA
        If IBERIA = True Then
            strCad = strCad & " AND {clientes.IBERIA} = 1 "
        End If
        ' vencimiento
        If vencimiento = True Then
            strCad = strCad & " AND {docs_pago.FP_ID} <> 0 "
'            strCad = strCad & " AND DATE_ADD(d.FECHA_FACTURA ,INTERVAL fp.dias DAY) >= '" & Format(fdesdev, "yyyy-mm-dd") & "'"
'            strCad = strCad & " AND DATE_ADD(d.FECHA_FACTURA ,INTERVAL fp.dias DAY) <= '" & Format(fhastav, "yyyy-mm-dd") & "'"
        Else
            ' Fdesde y fhasta
            strCad = strCad & " AND {docs_pago.FECHA_FACTURA} in Date (" & Format(fdesde, "yyyy, mm, dd") & ") to Date (" & Format(fhasta, "yyyy, mm, dd") & ")"
        End If
        'Enviada
        strCad = strCad & " AND {docs_pago.ENVIADA} in [" & ENVIADA & "]"
        
        
        ReDim arrNom(2)
        ReDim arrVal(2)
        
        arrNom(1) = "fecha_desde"
        arrVal(1) = Format(fdesde, "dd/mm/yyyy")
        
        arrNom(2) = "fecha_hasta"
        arrVal(2) = Format(fhasta, "dd/mm/yyyy")
        
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal
        
        .criterio = strCad
        .imprimir = False
        .generar
        .visible = True
        '.Show 1
    End With
    Set objfrm = Nothing

   On Error GoTo 0
   Exit Function

ImprimirListado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImprimirListado of Módulo de clase clsDocs_pago"

End Function

Public Property Let setID_DOC(ByVal dato As Long)
    ID_DOC = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
    NUMERO = dato
End Property
Public Property Let setTIPO(ByVal dato As Integer)
    TIPO = dato
End Property
Public Property Let setFECHA_FACTURA(ByVal dato As String)
    FECHA_FACTURA = Trim(dato)
End Property
Public Property Let setFECHA_GENERACION(ByVal dato As String)
    FECHA_GENERACION = Trim(dato)
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
    EMPLEADO_ID = Trim(dato)
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
    CLIENTE_ID = dato
End Property
Public Property Let setCLIENTE_ID_FACTURA(ByVal dato As Long)
    CLIENTE_ID_FACTURA = dato
End Property
Public Property Let setTOTAL(ByVal dato As String)
    total = Trim(dato)
End Property
Public Property Let setDESCUENTO(ByVal dato As String)
    DESCUENTO = dato
End Property
Public Property Let setIVA(ByVal dato As Integer)
    IVA = dato
End Property
Public Property Let setPAGADO(ByVal dato As Long)
    pagado = dato
End Property
Public Property Let setFACTURA_CONCEPTOS(ByVal dato As Integer)
    factura_conceptos = dato
End Property
Public Property Let setFP_ID(ByVal dato As Integer)
    FP_ID = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Integer)
    PEDIDO_ID = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
Public Property Let setENVIADA(ByVal dato As Integer)
    ENVIADA = dato
End Property
Public Property Let setASIENTO(ByVal dato As Integer)
    asiento = dato
End Property
Public Property Let setCONTABILIZADA(ByVal dato As String)
    CONTABILIZADA = dato
End Property
Public Property Let setFACTORIZADA(ByVal dato As String)
    FACTORIZADA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
    COMENTARIO = dato
End Property
Public Property Get getID_DOC() As Long
    getID_DOC = ID_DOC
End Property
Public Property Get getNUMERO() As Long
    getNUMERO = NUMERO
End Property
Public Property Get getTIPO() As Integer
    getTIPO = TIPO
End Property
Public Property Get getFECHA_FACTURA() As String
    getFECHA_FACTURA = FECHA_FACTURA
End Property
Public Property Get getFECHA_GENERACION() As String
    getFECHA_GENERACION = FECHA_GENERACION
End Property
Public Property Get getEMPLEADO_ID() As Integer
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getCLIENTE_ID() As String
    getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getCLIENTE_ID_FACTURA() As String
    getCLIENTE_ID_FACTURA = CLIENTE_ID_FACTURA
End Property
Public Property Get getTOTAL() As String
    getTOTAL = total
End Property
Public Property Get getDESCUENTO() As String
    getDESCUENTO = DESCUENTO
End Property
Public Property Get getIVA() As Integer
    getIVA = IVA
End Property
Public Property Get getPAGADO() As Long
    getPAGADO = pagado
End Property
Public Property Get getFACTURA_CONCEPTOS() As Integer
    getFACTURA_CONCEPTOS = factura_conceptos
End Property
Public Property Get getFP_ID() As Integer
    getFP_ID = FP_ID
End Property
Public Property Get getPEDIDO_ID() As Integer
    getPEDIDO_ID = PEDIDO_ID
End Property

Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Property Get getENVIADA() As Integer
    getENVIADA = ENVIADA
End Property
Public Property Get getASIENTO() As Long
    getASIENTO = asiento
End Property
Public Property Get getCONTABILIZADA() As String
    getCONTABILIZADA = CONTABILIZADA
End Property
Public Property Get getFACTORIZADA() As String
    getFACTORIZADA = FACTORIZADA
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getCOMENTARIO() As String
    getCOMENTARIO = COMENTARIO
End Property
Public Sub CrearId_doc()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT max(id_doc) FROM docs_pago"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_DOC = 1
    Else
        ID_DOC = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Sub CrearNumero(lTIPO As Integer)
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT max(numero) " & _
               "  FROM docs_pago " & _
               " WHERE tipo=" & lTIPO & " AND YEAR(FECHA_FACTURA)=" & Year(Date)
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
'        If ltipo = 1 Then
'            NUMERO = 5000 ' Albaranes empiezan en el 5000
'        Else
            NUMERO = 1
'        End If
    Else
        NUMERO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Private Function calcularIVA(TIPO As Integer, CLIENTE_ID As Long) As Integer
   On Error GoTo calcularIVA_Error
    Dim IVA As Integer
    If TIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN Then
        IVA = 0
    Else
        IVA = recuperaIVA()
        Dim oCliente As New clsCliente
        If oCliente.CargaCliente(CLIENTE_ID) Then
            If oCliente.getINTRA = 1 Or oCliente.getEXTRANJERO = 1 Then
                IVA = 0
            End If
        End If
    End If
    calcularIVA = IVA
   On Error GoTo 0
   Exit Function

calcularIVA_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularIVA of Módulo de clase clsDocs_pago"

End Function
Public Function InsertarDocPago() As Long
    Dim consulta As String
    On Error GoTo fallo:
    Me.CrearId_doc
    Me.CrearNumero (Me.getTIPO)
    Me.setCONTABILIZADA = 0
    Me.setANULADO = 0
    Me.setASIENTO = 0
    Me.setFACTORIZADA = ""
    Me.setENVIADA = 0
    Me.setIVA = calcularIVA(TIPO, CLIENTE_ID)
    ' Al insertar el doc_pago, el cliente de la factura será el mismo
    ' que el cliente de las muestras
    consulta = "INSERT INTO docs_pago VALUES (" & _
               ID_DOC & "," & TIPO & "," & NUMERO & ",'" & _
               FECHA_FACTURA & "',current_Date," & _
               EMPLEADO_ID & "," & CLIENTE_ID & "," & CLIENTE_ID_FACTURA & ",'" & _
               total & "','" & DESCUENTO & "'," & IVA & "," & pagado & "," & factura_conceptos & "," & _
               FP_ID & "," & PEDIDO_ID & ",'" & CONTABILIZADA & "','" & FACTORIZADA & "','" & _
               OBSERVACIONES & "','" & COMENTARIO & "'," & anulado & "," & ENVIADA & "," & asiento & ")"
    execute_bd consulta
    
    Dim oParametro As New clsParametros
    oParametro.Carga PARAM_USUARIO_VIGILADO, ""
    If USUARIO.getID_EMPLEADO = oParametro.getVALOR Then
        Dim ASUNTO As String
        Dim DETALLE As String
        Select Case TIPO
        Case 1
            ASUNTO = "El usuario " & USUARIO.getUSUARIO & " ha dado de alta un nuevo ALBARAN en geslab."
        Case 2
            ASUNTO = "El usuario " & USUARIO.getUSUARIO & " ha dado de alta una nueva FACTURA en geslab."
        Case 3
            ASUNTO = "El usuario " & USUARIO.getUSUARIO & " ha dado de alta un nuevo ABONO en geslab."
        Case 4
            ASUNTO = "El usuario " & USUARIO.getUSUARIO & " ha dado de alta una nueva PROFORMA en geslab."
        End Select
        
        DETALLE = "" & vbNewLine
        DETALLE = DETALLE & " Fecha : " & Format(Date, "dd-mm-yyyy") & vbNewLine
        DETALLE = DETALLE & " Hora  : " & Time & vbNewLine & vbNewLine
        DETALLE = DETALLE & " Número Factura : " & NUMERO & "/" & Year(FECHA_FACTURA) & vbNewLine
        Dim oCliente As New clsCliente
        oCliente.CargaCliente CLIENTE_ID
        DETALLE = DETALLE & " Cliente : " & oCliente.getNOMBRE & vbNewLine
        DETALLE = DETALLE & " Importe : " & total & vbNewLine
        oParametro.Carga PARAM_USUARIO_VIGILADO_CORREO, ""
        ret = Enviar_Mail_CDO(oParametro.getVALOR, ASUNTO, DETALLE, vbNullString)
    End If
    
    InsertarDocPago = ID_DOC
    Exit Function
fallo:
    InsertarDocPago = 0
    MsgBox "Error al guardar el documento de pago (InsertarDocPago) : " & Err.Description, vbCritical, Err.Description
End Function

Public Function CargarDocumento(num_doc As Long) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM docs_pago WHERE ID_DOC=" & num_doc
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            CargarDocumento = False
        Else
            ID_DOC = rs.Fields("ID_DOC")
            NUMERO = rs.Fields("NUMERO")
            TIPO = rs.Fields("TIPO")
            FECHA_FACTURA = rs.Fields("FECHA_FACTURA") & ""
            FECHA_GENERACION = rs.Fields("FECHA_GENERACION") & ""
            EMPLEADO_ID = rs.Fields("EMPLEADO_ID") & ""
            CLIENTE_ID = rs.Fields("CLIENTE_ID")
            CLIENTE_ID_FACTURA = rs.Fields("CLIENTE_ID_FACTURA")
            total = rs.Fields("TOTAL") & ""
            DESCUENTO = rs.Fields("DESCUENTO") & ""
            If IsNull(rs.Fields("IVA")) Then
               IVA = 0
            Else
               IVA = rs.Fields("IVA")
            End If
            If IsNull(rs.Fields("PAGADO")) Then
               pagado = 0
            Else
               pagado = rs.Fields("PAGADO")
            End If
            If IsNull(rs.Fields("ANULADO")) Then
               anulado = 0
            Else
               anulado = rs.Fields("ANULADO")
            End If
            If IsNull(rs.Fields("ENVIADA")) Then
               ENVIADA = 0
            Else
               ENVIADA = rs.Fields("ENVIADA")
            End If
            asiento = rs("ASIENTO")
            PEDIDO_ID = rs("PEDIDO_ID")
            FP_ID = rs("FP_ID")
            CONTABILIZADA = rs("CONTABILIZADA")
            factura_conceptos = rs.Fields("FACTURA_CONCEPTOS")
            FACTORIZADA = rs.Fields("FACTORIZADA")
            OBSERVACIONES = rs.Fields("OBSERVACIONES")
            COMENTARIO = rs("COMENTARIO")
            CargarDocumento = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargarDocumento = False
        MsgBox "Error al cargar el documento (CargaDocumento)", vbCritical, Err.Description
End Function
Public Function Modificar(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE docs_pago " & _
               "   SET descuento='" & DESCUENTO & "'," & _
               "       CLIENTE_ID=" & CLIENTE_ID & "," & _
               "       IVA=" & IVA & "," & _
               "       FP_ID=" & FP_ID & "," & _
               "       PEDIDO_ID=" & PEDIDO_ID & "," & _
               "       FECHA_FACTURA='" & FECHA_FACTURA & "'" & _
               " WHERE id_doc=" & documento
    execute_bd consulta
    Informar_total_factura (documento)
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el documento de pago.", vbCritical, Err.Description
End Function
Public Function Modificar_Factura_Conceptos(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE docs_pago " & _
               "   SET descuento='" & DESCUENTO & "'," & _
               "       CLIENTE_ID=" & CLIENTE_ID & "," & _
               "       CLIENTE_ID_FACTURA=" & CLIENTE_ID_FACTURA & "," & _
               "       IVA=" & IVA & "," & _
               "       FP_ID=" & FP_ID & "," & _
               "       PEDIDO_ID=" & PEDIDO_ID & "," & _
               "       FECHA_FACTURA='" & FECHA_FACTURA & "'" & _
               " WHERE id_doc=" & documento
    execute_bd consulta
    Informar_total_factura (documento)
    Modificar_Factura_Conceptos = True
    Exit Function
fallo:
    Modificar_Factura_Conceptos = False
    MsgBox "Error al modificar el documento de pago (Modificar_Factura_Conceptos).", vbCritical, Err.Description
End Function
Public Function Modificar_Factura_Muestras(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE docs_pago " & _
               "   SET descuento='" & DESCUENTO & "'," & _
               "       CLIENTE_ID=" & CLIENTE_ID & "," & _
               "       CLIENTE_ID_FACTURA=" & CLIENTE_ID_FACTURA & "," & _
               "       IVA=" & IVA & "," & _
               "       FP_ID=" & FP_ID & "," & _
               "       PEDIDO_ID=" & PEDIDO_ID & "," & _
               "       FECHA_FACTURA='" & FECHA_FACTURA & "'" & _
               " WHERE id_doc=" & documento
    execute_bd consulta
    Informar_total_factura (documento)
    Modificar_Factura_Muestras = True
    Exit Function
fallo:
    Modificar_Factura_Muestras = False
    MsgBox "Error al Modificar_Factura_Muestras en el documento de pago.", vbCritical, Err.Description
End Function

Public Function Modificar_Datos_Adicionales(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE docs_pago " & _
               "   SET descuento='" & DESCUENTO & "'," & _
               "       CLIENTE_ID=" & CLIENTE_ID & "," & _
               "       FP_ID=" & FP_ID & "," & _
               "       PEDIDO_ID=" & PEDIDO_ID & "," & _
               "       FECHA_FACTURA='" & FECHA_FACTURA & "'" & _
               " WHERE id_doc=" & documento
    execute_bd consulta
    Informar_total_factura (documento)
    Modificar_Datos_Adicionales = True
    Exit Function
fallo:
    Modificar_Datos_Adicionales = False
    MsgBox "Error al modificar el documento de pago.", vbCritical, Err.Description
End Function
Public Function informar_factura_conceptos(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    Dim muestras As Boolean
    Dim conceptos As Boolean
    Dim rs As ADODB.Recordset
    consulta = "SELECT DOC_ID FROM DOCS_PAGO_MUESTRAS WHERE DOC_ID = " & documento
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        muestras = True
    Else
        muestras = False
    End If
    consulta = "SELECT DOC_ID FROM DOCS_PAGO_CONCEPTOS WHERE DOC_ID = " & documento
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        conceptos = True
    Else
        conceptos = False
    End If
    Dim TIPO As Integer
    If muestras And conceptos Then
        TIPO = 2
    ElseIf muestras Then
        TIPO = 0
    Else
        TIPO = 1
    End If
    consulta = "UPDATE docs_pago " & _
               "   SET factura_conceptos=" & TIPO & _
               " WHERE id_doc=" & documento
    execute_bd consulta
    informar_factura_conceptos = True
    Exit Function
fallo:
    informar_factura_conceptos = False
    MsgBox "Error al modificar el documento de pago.", vbCritical, Err.Description
End Function
Public Function Informar_total_factura(documento As Long) As Boolean
    On Error GoTo fallo
'    Dim odoc As New clsDocs_pago
    Dim IMPORTE As Currency
'    IMPORTE = odoc.ImporteTotalDocumento(CLng(documento))
    IMPORTE = ImporteTotalDocumento(CLng(documento))
    execute_bd "update docs_pago set total = '" & Replace(Format(IMPORTE, "0.00"), ",", ".") & "' where id_doc = " & documento
    Informar_total_factura = True
    Exit Function
fallo:
    Informar_total_factura = False
    MsgBox "Error al modificar el total del documento de pago.", vbCritical, Err.Description
End Function
Public Function ClientesConDocumentosPago() As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT DISTINCT A.ID_CLIENTE,A.NOMBRE " & _
               "  FROM CLIENTES A,DOCS_PAGO B" & _
               " WHERE A.ANULADO=0 " & _
               "   AND A.ID_CLIENTE = B.CLIENTE_ID " & _
               " ORDER BY A.NOMBRE "
'    consulta = "SELECT id_cliente,nombre FROM clientes where anulado=0 ORDER BY nombre "
    Set ClientesConDocumentosPago = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'ClientesConDocumentosPago'. ", vbCritical, Err.Description
End Function
Public Function ClientesConAlbaranesPendientes() As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT DISTINCT A.ID_CLIENTE,A.NOMBRE " & _
               "  FROM CLIENTES A,DOCS_PAGO B" & _
               " WHERE B.TIPO=1 AND A.ANULADO=0 AND B.PAGADO=0 AND B.ANULADO = 0 " & _
               "   AND A.ID_CLIENTE = B.CLIENTE_ID " & _
               " ORDER BY A.nombre "
    Set ClientesConAlbaranesPendientes = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'ClientesConDocumentosPago'. ", vbCritical, Err.Description
End Function
Public Function Documento_por_numero(NUMERO As Long, ANNO As Integer, lTIPO As String, lpagado As Integer, ESTADO As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim finicial As String
    Dim ffinal As String
    Dim TIPO As String
    finicial = Format(Trim(str(ANNO)) & "-01-01", "yyyy-mm-dd")
    ffinal = Format(Trim(str(ANNO)) & "-12-31", "yyyy-mm-dd")
'    If lTIPO <> 0 Then
'        tipo = " AND d.TIPO=" & lTIPO
'    Else
'        tipo = " AND d.TIPO in (2,3) "
'    End If
    If lTIPO <> "" Then
        TIPO = " AND d.TIPO in (" & lTIPO & ") "
    End If
    consulta = "SELECT d.ID_DOC,d.NUMERO,cl.NOMBRE,d.FECHA_FACTURA,d.DESCUENTO,d.IVA,d.TIPO,d.PAGADO," & _
               "       d.total,d.FACTURA_CONCEPTOS,d.PEDIDO_ID,d.FACTORIZADA,IFNULL(cp.codigo,''),d.FP_ID, " & _
               "       d.CLIENTE_ID,d.CLIENTE_ID_FACTURA,d.asiento, fp.DIAS, d.COMENTARIO, clifact.nombre, dpc.fecha,clifact.CC,clifact.CIF,dpc.FECHA_PREVISTA  " & _
               "  FROM DOCS_PAGO d " & _
               " INNER JOIN clientes as cl ON cl.id_cliente = d.cliente_id" & _
               " LEFT JOIN clientes_pedidos cp ON d.pedido_id = cp.id_pedido " & _
               " INNER JOIN formas_pago fp ON d.FP_ID = fp.ID_FP " & _
               " INNER JOIN clientes as clifact ON d.cliente_id_factura = clifact.id_cliente " & _
               " LEFT JOIN docs_pago_cobros as dpc ON d.id_doc = dpc.doc_id " & _
               " WHERE 1=1 " & TIPO & _
               "   AND d.numero = " & NUMERO & _
               "   AND FECHA_FACTURA BETWEEN '" & finicial & "' AND '" & ffinal & "' " & _
               "   AND d.ANULADO in (" & ESTADO & ")" & _
               " ORDER BY d.id_doc desc"
    Set Documento_por_numero = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'Documento_por_numero'. ", vbCritical, Err.Description
End Function

Public Function Albaranes(cliente As Long, AIRBUS As Boolean, IBERIA As Boolean, fdesde As String, fhasta As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim s As String
    If cliente <> 0 Then
        s = s & " AND cl.id_cliente = " & cliente
    End If
    If AIRBUS = True Then
        s = s & " AND cl.AIRBUS = 1"
    End If
    If IBERIA = True Then
        s = s & " AND cl.IBERIA = 1"
    End If
    If fdesde <> "" Then
        s = s & " AND d.FECHA_FACTURA >='" & Format(fdesde, "yyyy-mm-dd") & "'"
    End If
    If fhasta <> "" Then
        s = s & " AND d.FECHA_FACTURA <='" & Format(fhasta, "yyyy-mm-dd") & "'"
    End If
    consulta = "SELECT d.ID_DOC,d.NUMERO,cl.NOMBRE,d.FECHA_FACTURA,d.DESCUENTO,d.IVA,d.TIPO,d.PAGADO,d.total,d.FACTURA_CONCEPTOS,d.PEDIDO_ID,cp.DESCRIPCION,d.CLIENTE_ID  " & _
               "  FROM docs_pago d " & _
               " INNER JOIN clientes AS cl ON cl.id_cliente = d.cliente_id " & _
               "  LEFT JOIN clientes_pedidos as cp ON d.PEDIDO_ID = cp.ID_PEDIDO " & _
               " WHERE d.ANULADO = 0 AND d.TIPO=1 " & _
               "   AND d.PAGADO = 0 " & _
               s & _
               " order by d.pedido_id,d.id_doc desc"
    Set Albaranes = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'Albaranes'. ", vbCritical, Err.Description
End Function

Public Function Documentos(cliente As Long, clienteFact As Long, fdesde As String, fhasta As String, lTIPO As String, lpagado As Integer, ESTADO As String, ENVIADA As String, AIRBUS As Boolean, vencimiento As Boolean, fdesdev As String, fhastav As String, fcobro As Boolean, fCobroDesde As String, fCobroHasta As String, fPrevista As Boolean, fPrevistaDesde As String, fPrevistaHasta As String, fp As Long, IBERIA As Boolean, concepto As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim TIPO As String
    Dim join As String
    
'    If lTIPO <> 0 Then
'        tipo = " AND d.TIPO=" & lTIPO
    If lTIPO <> "" Then
        TIPO = " AND d.TIPO in (" & lTIPO & ") "
    Else
        If lpagado = 0 Then
            TIPO = " AND d.TIPO = 2 "
        Else
            TIPO = " AND d.TIPO in (2,3) "
        End If
    End If
    If Trim(concepto) <> "" Then
        join = " INNER JOIN docs_pago_conceptos dpcon ON d.id_doc = dpcon.doc_id "
    End If
    consulta = "SELECT DISTINCT d.ID_DOC,d.NUMERO,cl.NOMBRE,d.FECHA_FACTURA, " & _
               "       d.DESCUENTO,d.IVA,d.TIPO,d.PAGADO,ifnull(d.total,0), " & _
               "       d.FACTURA_CONCEPTOS,d.PEDIDO_ID,d.FACTORIZADA, " & _
               "       IFNULL(cp.CODIGO,''),d.FP_ID,d.CLIENTE_ID,d.CLIENTE_ID_FACTURA,d.asiento, fp.DIAS, d.COMENTARIO, clifact.nombre,ifnull(dpc.FECHA,''),clifact.CC,clifact.CIF,ifnull(dpc.FECHA_PREVISTA,'') " & _
               "  FROM docs_pago d " & _
               " INNER JOIN clientes as cl ON cl.id_cliente = d.cliente_id " & _
               " LEFT JOIN clientes_pedidos cp ON d.pedido_id = cp.id_pedido " & _
               " INNER JOIN formas_pago fp ON d.FP_ID = fp.ID_FP " & _
               " INNER JOIN clientes as clifact ON d.cliente_id_factura = clifact.id_cliente " & _
               "  LEFT JOIN docs_pago_cobros as dpc ON d.id_doc = dpc.doc_id " & _
               join & _
               " WHERE d.ANULADO in (" & ESTADO & ") " & TIPO
    If Trim(concepto) <> "" Then
        consulta = consulta & " AND dpcon.descripcion like '%" & Trim(concepto) & "%'"
    End If
    
    ' LPAGADO = PENDIENTE DE COBRO
    ' 0 : Si
    ' 1 : No
    ' 2 : Todos
    If lpagado < 2 Then
        consulta = consulta & " AND d.PAGADO= " & lpagado
    End If
    If AIRBUS = True Then
        consulta = consulta & " AND cl.AIRBUS = 1 "
    End If
    If IBERIA = True Then
        consulta = consulta & " AND cl.IBERIA = 1 "
    End If
    If vencimiento = True Then
        consulta = consulta & " AND d.FP_ID <> 0 "
        consulta = consulta & " AND DATE_ADD(d.FECHA_FACTURA ,INTERVAL fp.dias DAY) >= '" & Format(fdesdev, "yyyy-mm-dd") & "'"
        consulta = consulta & " AND DATE_ADD(d.FECHA_FACTURA ,INTERVAL fp.dias DAY) <= '" & Format(fhastav, "yyyy-mm-dd") & "'"
    Else
        consulta = consulta & " AND d.FECHA_FACTURA >='" & fdesde & "' "
        consulta = consulta & " AND d.FECHA_FACTURA <='" & fhasta & "' "
    End If
    If fcobro = True Then
        consulta = consulta & " AND dpc.fecha >='" & Format(fCobroDesde, "yyyy-mm-dd") & "'"
        consulta = consulta & " AND dpc.fecha <='" & Format(fCobroHasta, "yyyy-mm-dd") & "'"
    End If
    If fPrevista = True Then
        consulta = consulta & " AND dpc.fecha_prevista >='" & Format(fPrevistaDesde, "yyyy-mm-dd") & "'"
        consulta = consulta & " AND dpc.fecha_prevista <='" & Format(fPrevistaHasta, "yyyy-mm-dd") & "'"
    End If
    If cliente <> 0 Then
        consulta = consulta & " AND d.CLIENTE_ID =  " & cliente
    End If
    If clienteFact <> 0 Then
        consulta = consulta & " AND d.CLIENTE_ID_FACTURA =  " & clienteFact
    End If
    If fp <> 0 Then
        consulta = consulta & " AND d.FP_ID =  " & fp
    End If
    consulta = consulta & " AND d.ENVIADA in (" & ENVIADA & ") "
    consulta = consulta & " order by id_doc desc"
    Set Documentos = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'Documentos'. ", vbCritical, Err.Description
End Function


Public Function Documentos_por_pedido(pedido As Long) As ADODB.Recordset
    Dim consulta As String
   On Error GoTo Documentos_por_pedido_Error

    consulta = "SELECT d.ID_DOC,d.NUMERO,cl.NOMBRE,d.FECHA_FACTURA,d.DESCUENTO,d.IVA,d.TIPO,d.PAGADO,d.TOTAL,d.FACTURA_CONCEPTOS,d.FP_ID " & _
               "  FROM docs_pago d, clientes as cl " & _
               " WHERE d.ANULADO = 0 " & _
               "   AND cl.id_cliente = d.cliente_id " & _
               "   AND d.pedido_id = " & pedido & _
               "   AND d.tipo = 2 " & _
               " ORDER BY id_doc desc"
    Set Documentos_por_pedido = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

Documentos_por_pedido_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Documentos_por_pedido of Módulo de clase clsDocs_pago"
End Function
'Public Function Documentos_por_pedido_suma_importe(pedido As Long) As Currency
'    Dim consulta As String
'   On Error GoTo Documentos_por_pedido_Error
'
'    consulta = "SELECT sum(d.TOTAL) " & _
'               "  FROM docs_pago d, clientes as cl " & _
'               "WHERE d.ANULADO = 0 " & _
'               "  AND cl.id_cliente = d.cliente_id " & _
'               "  AND d.tipo = 2 " & _
'               "  AND d.pedido_id = " & pedido
'    Dim rs As ADODB.Recordset
'    Set rs = datos_bd(consulta)
'    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
'        Documentos_por_pedido_suma_importe = moneda(0)
'    Else
'        Documentos_por_pedido_suma_importe = moneda(rs.Fields(0))
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'Documentos_por_pedido_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Documentos_por_pedido_suma_importe of Módulo de clase clsDocs_pago"
'End Function

Public Function Cobrar(documento As Long) As Boolean
    On Error GoTo fallo
    Dim oDP As New clsDocs_pago
    oDP.CargarDocumento documento
    If oDP.getTIPO <> C_TIPOS_DOCS_PAGO_ALBARAN Then
        Dim consulta As String
        consulta = "UPDATE docs_pago SET pagado = 1 where id_doc =" & documento
        execute_bd consulta
    End If
    Cobrar = True
    Exit Function
fallo:
    Cobrar = False
    MsgBox "Error al marcar como cobrado el documento.", vbCritical, Err.Description
End Function
Public Function DesCobrar(documento As Long) As Boolean
    On Error GoTo fallo
    Dim oDP As New clsDocs_pago
    oDP.CargarDocumento documento
    If oDP.getTIPO <> C_TIPOS_DOCS_PAGO_ALBARAN Then
        Dim consulta As String
        consulta = "UPDATE docs_pago SET pagado = 0 where id_doc =" & documento
        execute_bd consulta
        consulta = "UPDATE docs_pago_cobros SET fecha = null, hora = null where doc_id =" & documento
        execute_bd consulta
    End If
    DesCobrar = True
    Exit Function
fallo:
    DesCobrar = False
    MsgBox "Error al marcar como no cobrado el documento.", vbCritical, Err.Description
End Function
Public Function Anular(documento As Long, OBSERVACIONES As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE docs_pago " & _
               " SET anulado = 1, " & _
               "     OBSERVACIONES = '" & OBSERVACIONES & "' " & _
               " where id_doc =" & documento
    execute_bd consulta
    consulta = "UPDATE alodine_planificacion SET doc_id = 0 where doc_id =" & documento
    execute_bd consulta
    consulta = "UPDATE suministros_lotes_clientes SET doc_id = 0 where doc_id =" & documento
    execute_bd consulta
    Anular = True
    Exit Function
fallo:
    Anular = False
    MsgBox "Error al anular el documento.", vbCritical, Err.Description
End Function
Public Function Abonar(documento As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE docs_pago SET anulado = 2 where id_doc =" & documento
    execute_bd consulta
    Abonar = True
    Exit Function
fallo:
    Abonar = False
    MsgBox "Error al abonar el documento.", vbCritical, Err.Description
End Function
Public Function generar_factura(ID As Long, imprimir As Boolean, generar_pdf As String, PLANTILLA As String) As Boolean
    Dim oDoc As New clsDocs_pago
    oDoc.CargarDocumento ID
    Dim oCliente As New clsCliente
    oCliente.CargaCliente oDoc.getCLIENTE_ID_FACTURA
    If oCliente.getIDIOMA_FACTURA = 1 Then 'Ingles
        PLANTILLA = PLANTILLA & "_EN"
    End If
    ' Parametros del documento de pago (LOGO)
    With frmReport
        .iniciar
        .informe = "\Facturacion\" & PLANTILLA
        .criterio = "{docs_pago.ID_DOC}=" & CStr(ID)
        .imprimir = imprimir
        .pdf = generar_pdf
        If imprimir = True Or generar_pdf <> "" Then
            .generar
            .visible = False
        Else
            .generar
            .Show 1
        End If
    End With
    Unload frmReport
End Function
Public Function Otros_datos_documento(ID As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT d.numero,c.nombre,d.fecha_factura,d.descuento,fp.nombre,p.codigo,d.id_doc,d.fp_id,d.pedido_id,d.cliente_id " & _
               "  FROM docs_pago d,clientes c,formas_pago fp,clientes_pedidos p " & _
               " WHERE d.cliente_id = c.id_cliente " & _
               "   AND d.fp_id = fp.id_fp " & _
               "   AND d.pedido_id = p.id_pedido " & _
               "   AND d.id_doc = " & ID
    Set Otros_datos_documento = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'ClientesConDocumentosPago'. ", vbCritical, Err.Description
End Function

Private Function ImporteTotalDocumento(num_documento As Long) As Currency
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    ImporteTotalDocumento = 0
    'MDET
    consulta = " SELECT SUM(PRECIO) " & _
               "   FROM docs_pago_muestras " & _
               "  WHERE DOC_ID=" & num_documento & _
               "    AND MUESTRA_ID <> 0 AND DETERMINACION_ID = 0" & _
               "  UNION " & _
               " SELECT SUM(TOTAL) " & _
               "   FROM docs_pago_conceptos " & _
               "  WHERE DOC_ID=" & num_documento & _
               "    AND APARTADO = 0 "
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            If Not IsNull(rs.Fields(0)) And Not (rs.EOF And rs.BOF) Then
                ImporteTotalDocumento = ImporteTotalDocumento + rs.Fields(0)
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Exit Function
fallo:
    ImporteTotalDocumento = 0
    MsgBox "Error al calcular el importe total del documento (ImporteDocumento)", vbCritical, Err.Description
End Function
Public Function Facturar_Albaran(ID As Long, factura As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE DOCS_PAGO " & _
               "   SET PAGADO=" & factura & _
               " WHERE ID_DOC=" & ID
    execute_bd consulta
    Facturar_Albaran = True
    Exit Function
fallo:
    Facturar_Albaran = False
    MsgBox "Error al Facturar_Albaran.", vbCritical, Err.Description
End Function

Public Function Numero_Cartas_Pago(criterio As String) As Integer
     On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT COUNT(*)  " & _
               "  FROM DOCS_PAGO " & _
               " WHERE ID_DOC IN (" & criterio & ")" & _
               " GROUP BY CLIENTE_ID"
    Set rs = datos_bd(consulta)
    Numero_Cartas_Pago = rs.RecordCount
    Exit Function
fallo:
    Numero_Cartas_Pago = 0
    MsgBox "Error al calcular el numero de cartas de pago. ", vbCritical, Err.Description
End Function
Public Function Clientes_Distintos(criterio As String) As ADODB.Recordset
     On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT CLIENTE_ID  " & _
               "  FROM DOCS_PAGO " & _
               " WHERE ID_DOC IN (" & criterio & ")" & _
               " GROUP BY CLIENTE_ID"
    Set Clientes_Distintos = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al calcular el numero de clientes ocn cartas de pago. ", vbCritical, Err.Description
End Function
Public Function Facturas_Pendientes_Cliente(criterio As String, cliente As Long) As ADODB.Recordset
     On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT ID_DOC,NUMERO,FECHA_FACTURA  " & _
               "  FROM DOCS_PAGO " & _
               " WHERE ID_DOC IN (" & criterio & ")" & _
               "   AND CLIENTE_ID = " & cliente & _
               " GROUP BY ID_DOC ASC"
    Set Facturas_Pendientes_Cliente = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al calcular el numero de clientes ocn cartas de pago. ", vbCritical, Err.Description
End Function

Public Function Listado_contabilidad(desde As Date, hasta As Date, contabilizado As Boolean, NUMERO As String, ANNO As String) As ADODB.Recordset
   On Error GoTo Listado_contabilidad_Error

    On Error GoTo fallo
    Dim consulta As String
    Dim conta As String
    If contabilizado = True Then
        conta = conta & "   AND d.contabilizada <> '0' "
    Else
        conta = conta & "   AND d.contabilizada = '0' "
    End If
    If NUMERO <> "" Then
        If IsNumeric(NUMERO) Then
            conta = conta & "    AND d.numero = " & NUMERO & " and year(d.fecha_factura) = " & ANNO
        End If
    Else
        conta = conta & "   AND d.fecha_factura >= '" & Format(desde, "yyyy-mm-dd") & "'"
        conta = conta & "   AND d.fecha_factura <= '" & Format(hasta, "yyyy-mm-dd") & "'"
    End If
    consulta = "SELECT d.ID_DOC,d.NUMERO,cl.NOMBRE,d.FECHA_FACTURA,d.DESCUENTO,d.IVA,d.TIPO,d.PAGADO,ifnull(d.total,0),d.FACTURA_CONCEPTOS,d.PEDIDO_ID, cl.CC  " & _
               "  FROM docs_pago d, clientes as cl " & _
               " WHERE (d.ANULADO = 0 or d.ANULADO = 2) AND (d.TIPO=2 OR d.TIPO=3) " & _
               "   AND cl.id_cliente = d.cliente_id " & _
               conta & _
               " order by d.id_doc asc"
    Set Listado_contabilidad = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'Listado_contabilidad'. ", vbCritical, Err.Description

   On Error GoTo 0
   Exit Function

Listado_contabilidad_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Listado_contabilidad of Módulo de clase clsDocs_pago"
End Function
Public Function contabilizar(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    Dim fecha As String
    fecha = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
    consulta = "UPDATE docs_pago SET CONTABILIZADA='" & fecha & "' WHERE id_doc=" & documento
    execute_bd consulta
    contabilizar = True
    Exit Function
fallo:
    contabilizar = False
    MsgBox "Error al contabilizar el documento de pago.", vbCritical, Err.Description
End Function
Public Function comentar(documento As Long, comen As String) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE docs_pago SET COMENTARIO='" & comen & "' WHERE id_doc=" & documento
    execute_bd consulta
    comentar = True
    Exit Function
fallo:
    comentar = False
    MsgBox "Error al comentar el documento de pago.", vbCritical, Err.Description
End Function

Public Function no_contabilizar(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE docs_pago SET CONTABILIZADA='0' WHERE id_doc=" & documento
    execute_bd consulta
    no_contabilizar = True
    Exit Function
fallo:
    no_contabilizar = False
    MsgBox "Error al contabilizar el documento de pago.", vbCritical, Err.Description
End Function
'Public Function FACTORIZAR(documento As Long, fecha As String) As Boolean
'    Dim consulta As String
'    On Error GoTo fallo
'    consulta = "UPDATE docs_pago SET FACTORIZADA='" & fecha & "' WHERE id_doc=" & documento
'    execute_bd consulta
'    FACTORIZAR = True
'    Exit Function
'fallo:
'    FACTORIZAR = False
'    MsgBox "Error al FACTORIZAR el documento de pago.", vbCritical, Err.Description
'End Function
Public Function esta_contabilidado(num_doc As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT CONTABILIZADA FROM docs_pago WHERE ID_DOC=" & num_doc
    Set rs = datos_bd(consulta)
    If rs(0) = "0" Then
        esta_contabilidado = False
    Else
        esta_contabilidado = True
        MsgBox "El documento ya se encuentra contabilizado. Recuerde contabilizarlo nuevamente si lo modifica.", vbCritical, App.Title
    End If
    Set rs = Nothing
    Exit Function
fallo:
    esta_contabilidado = False
    MsgBox "Error al verificar la contabilidad (CargaDocumento)", vbCritical, Err.Description
End Function

Public Function mensaje_abono(documento As Long, abono As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    Dim oDoc As New clsDocs_pago
    Dim oAbono As New clsDocs_pago
    Dim oB As String
    oAbono.CargarDocumento abono
    oDoc.CargarDocumento documento
    If oDoc.getOBSERVACIONES = "" Then
        oB = "FACTURA ABONADA EN : " & oAbono.getNUMERO & "/" & Format(oAbono.getFECHA_FACTURA, "yyyy")
    Else
        oB = oDoc.getOBSERVACIONES & "," & oAbono.getNUMERO & "/" & Format(oAbono.getFECHA_FACTURA, "yyyy")
    End If
    consulta = "UPDATE docs_pago " & _
               "   SET OBSERVACIONES='" & oB & "'" & _
               " WHERE id_doc=" & documento
    execute_bd consulta
    mensaje_abono = True
    Exit Function
fallo:
    mensaje_abono = False
    MsgBox "Error al insertar el mensaje de abono.", vbCritical, Err.Description
End Function
Public Function validar_previos_documento(DOC As Long) As Boolean
    Dim diferencia As Boolean
    Dim salida As String
    Dim sumatorio As Currency
    Dim precio_Analisis As Currency
    Dim descripcion_analisis As String
   On Error GoTo validar_previos_documento_Error
    validar_previos_documento = True
    diferencia = False
    sumatorio = 0
    precio_Analisis = 0
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocs_pago
    Dim dpm As New clsDocs_pago_muestras
    If oDoc.CargarDocumento(DOC) Then
        Set rs = dpm.lineas_factura(DOC)
        If rs.RecordCount > 0 Then
            salida = "Factura nº " & oDoc.getNUMERO & " Fecha : " & Format(oDoc.getFECHA_FACTURA, "dd-mm-yyyy") & vbNewLine
            salida = salida & "Revise los siguientes anális, el precio total no coincide con el de sus determinaciones : " & vbNewLine & vbNewLine
            While Not rs.EOF
'MDET
'                  If rs(0) <> 0 Then
                  If rs(8) = 0 Then
                    If precio_Analisis <> 0 And sumatorio <> 0 Then
                        If precio_Analisis <> sumatorio Then
                            diferencia = True
                            salida = salida & descripcion_analisis & vbNewLine
                        End If
                        sumatorio = 0
                    End If
                    precio_Analisis = CCur(rs(7))
                    descripcion_analisis = "* (" & rs(3) & ") " & rs(4) & " - " & rs(5)
                  Else
                    sumatorio = sumatorio + CCur(rs(7))
                  End If
                  rs.MoveNext
            Wend
            If diferencia Then
'                validar_previos_documento = False
                MsgBox salida
            End If
        End If
    End If
    Set rs = Nothing
    Set dpm = Nothing
    Set oDoc = Nothing
    
   On Error GoTo 0
   Exit Function

validar_previos_documento_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validar_previos_documento of Módulo de clase clsDocs_pago"
End Function

Public Function Enviar(documento As Long, ESTADO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE docs_pago SET ENVIADA = " & ESTADO & " where id_doc =" & documento
    execute_bd consulta
    Enviar = True
    Exit Function
fallo:
    Enviar = False
    MsgBox "Error al abonar el documento.", vbCritical, Err.Description
End Function
Public Function informar_asiento(documento As Long, lasiento As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE docs_pago SET ASIENTO = " & lasiento & " where id_doc =" & documento
    execute_bd consulta
    informar_asiento = True
    Exit Function
fallo:
    informar_asiento = False
    MsgBox "Error en el asiento del documento.", vbCritical, Err.Description
End Function
Public Function informar_pedido(ID_DOCUMENTO As Long, PEDIDO_ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE docs_pago SET PEDIDO_ID = " & PEDIDO_ID & " where ID_DOC =" & ID_DOCUMENTO
    execute_bd consulta
    informar_pedido = True
    Exit Function
fallo:
    informar_pedido = False
    MsgBox "Error en el metodo informar_pedido del documento.", vbCritical, Err.Description
End Function
Public Function MARCAR_FACTURADA(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE docs_pago SET PAGADO=1 WHERE id_doc=" & documento
    execute_bd consulta
    MARCAR_FACTURADA = True
    Exit Function
fallo:
    MARCAR_FACTURADA = False
    MsgBox "Error al MARCAR_FACTURADA el documento de pago.", vbCritical, Err.Description
End Function

Public Function recuperarIdPorNumero(NUMERO As Long, ANNO As Long) As Long
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_DOC FROM docs_pago " & _
               " WHERE NUMERO=" & NUMERO & _
               "   AND YEAR(FECHA_FACTURA) = " & ANNO & _
               "   AND TIPO = " & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        recuperarIdPorNumero = 0
    Else
        recuperarIdPorNumero = rs(0)
    End If
    Exit Function
fallo:
    recuperarIdPorNumero = 0
    MsgBox "Error al cargar el documento (CargaDocumento)", vbCritical, Err.Description
End Function

Public Function DocumentosTesoreria(FP_ID As Integer, TIPO As Integer, PERIODO As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim filtro As String
    Select Case TIPO
        Case 3 'AIRBUS SIN PEDIDO
            filtro = filtro & " and cl.airbus = 1 and d.pedido_id = 0 "
        Case 4 'AIRBUS CON PEDIDO
            filtro = filtro & " and cl.airbus = 1 and d.pedido_id <> 0 "
        Case 5 ' OTROS
            filtro = filtro & " and cl.airbus = 0 "
    End Select
    
    If PERIODO = "VENCIDAS" Then
        filtro = filtro & " AND d.FECHA_FACTURA < concat(year(current_date),'-',month(current_date),'-01') "
    Else
        filtro = filtro & " AND DATE_FORMAT(d.FECHA_FACTURA,'%m-%Y') = '" & PERIODO & "'"
    End If
    filtro = filtro & " AND d.FP_ID = " & FP_ID
    consulta = "SELECT d.ID_DOC,d.NUMERO,cl.NOMBRE,d.FECHA_FACTURA, " & _
               "       d.DESCUENTO,d.IVA,d.TIPO,d.PAGADO,d.total, " & _
               "       d.FACTURA_CONCEPTOS,d.PEDIDO_ID,d.FACTORIZADA, " & _
               "       cp.CODIGO,d.FP_ID,d.CLIENTE_ID,d.CLIENTE_ID_FACTURA,d.asiento, fp.DIAS, d.COMENTARIO, clifact.nombre,dpc.FECHA  " & _
               "  FROM docs_pago d " & _
               " INNER JOIN clientes as cl ON cl.id_cliente = d.cliente_id " & _
               " INNER JOIN clientes_pedidos cp ON d.pedido_id = cp.id_pedido " & _
               " INNER JOIN formas_pago fp ON d.FP_ID = fp.ID_FP " & _
               " INNER JOIN clientes as clifact ON d.cliente_id_factura = clifact.id_cliente " & _
               "  LEFT JOIN docs_pago_cobros as dpc ON d.id_doc = dpc.doc_id " & _
               " WHERE d.ANULADO = 0 and d.PAGADO = 0 and year(d.FECHA_FACTURA) >= 2010 " & _
               "   AND d.tipo in (" & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA & "," & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ABONO & ") " & _
                filtro & _
                " order by 1 desc"
    Set DocumentosTesoreria = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la clase 'Documentos'. ", vbCritical, Err.Description
End Function
Public Function VerificarAlbaranes(ID_DOC As Long) As String
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo VerificarAlbaranes_Error

    consulta = "SELECT DISTINCT ALBARAN_ID FROM DOCS_PAGO_CONCEPTOS " & _
               " WHERE DOC_ID = " & ID_DOC & _
               "   AND ALBARAN_ID <> 0"
    Set rs = datos_bd(consulta)
    Dim docs As String
    If rs.RecordCount > 0 Then
        Do
            Dim oDP As New clsDocs_pago
            oDP.CargarDocumento rs(0)
            If docs <> "" Then
                docs = docs & ","
            End If
            docs = docs & oDP.getNUMERO_FORMATEADO & "/" & Year(oDP.getFECHA_FACTURA)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    VerificarAlbaranes = docs

   On Error GoTo 0
   Exit Function

VerificarAlbaranes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VerificarAlbaranes of Módulo de clase clsDocs_pago"
    
End Function
Public Function ListadoDesgloseContable(ID_DOC As Long) As ADODB.Recordset
     On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT GROUP_CONCAT(DISTINCT ALBARAN_ID) FROM DOCS_PAGO_CONCEPTOS " & _
               " WHERE DOC_ID = " & ID_DOC & _
               "   AND ALBARAN_ID <> 0;"
    Set rs = datos_bd(consulta)
    Dim docs As String
    If IsNull(rs(0)) Then
        docs = ID_DOC
    Else
        docs = rs(0)
    End If
    'MDET
    consulta = "select c2, c3, sum(c4) from (" & _
        " select a.ID_DOC,a.NUMERO as c1,c.NOMBRE as c2,c.CC as c3,sum(b.TOTAL) as c4 from docs_pago a,docs_pago_conceptos b, familias c " & _
        " where a.ID_DOC = b.DOC_ID " & _
        " and b.FAMILIA_ID = c.ID_FAMILIA and b.apartado = 0 " & _
        " and a.id_doc IN (" & docs & ")" & _
        " group by 1,2,3,4 " & _
        " Union " & _
        " select a.ID_DOC,a.NUMERO as c1,c.NOMBRE as c2,c.CC as c3,sum(b.PRECIO)  as c4 " & _
        " from docs_pago a,docs_pago_muestras b, familias c, muestras d, tipos_muestra tm " & _
        " where a.ID_DOC = b.DOC_ID " & _
        " and b.MUESTRA_ID = d.ID_MUESTRA and b.DETERMINACION_ID = 0 " & _
        " and a.id_doc IN (" & docs & ")" & _
        " and d.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA " & _
        " and tm.FAMILIA_ID = c.ID_FAMILIA " & _
        " and b.muestra_id <> 0 " & _
        " group by 1,2,3,4 " & _
        " ) aa " & _
        " group by 1,2"
    Set ListadoDesgloseContable = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error ListadoDesgloseContable.", vbCritical, Err.Description
End Function


