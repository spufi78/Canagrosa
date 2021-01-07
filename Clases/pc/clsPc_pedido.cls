VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPc_pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPC_PEDIDO
'***************************************************************
Private ID_PEDIDO As Long
Private CLIENTE_ID As Long
Private TIPO_BOTE_EX_ID As Long
Private FECHA_PEDIDO As String
Private Identificacion As String
Private cantidad As Long
Private PEDIDO_CLIENTE As String
Private FECHA_CLIENTE As String
Private PRECIO As String
Private PEDIDO_ID As Long
Private DOC_ID As Long
Private NCERTIFICADO As Long
Private NORMATIVA_ID As Long
Private F_CADUCIDAD As String
Private IDENTIFICACION2 As String
Private INFORMES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPC_PEDIDO
'***************************************************************
Public Property Let setID_PEDIDO(ByVal dato As Long)
        ID_PEDIDO = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setTIPO_BOTE_EX_ID(ByVal dato As Long)
        TIPO_BOTE_EX_ID = dato
End Property
Public Property Let setFECHA_PEDIDO(ByVal dato As String)
        FECHA_PEDIDO = dato
End Property
Public Property Let setIDENTIFICACION(ByVal dato As String)
        Identificacion = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Long)
        cantidad = dato
End Property
Public Property Let setPEDIDO_CLIENTE(ByVal dato As String)
        PEDIDO_CLIENTE = dato
End Property
Public Property Let setFECHA_CLIENTE(ByVal dato As String)
        FECHA_CLIENTE = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setDOC_ID(ByVal dato As Long)
        DOC_ID = dato
End Property
Public Property Let setNCERTIFICADO(ByVal dato As Long)
        NCERTIFICADO = dato
End Property
Public Property Let setNORMATIVA_ID(ByVal dato As Long)
        NORMATIVA_ID = dato
End Property
Public Property Let setF_CADUCIDAD(ByVal dato As String)
        F_CADUCIDAD = dato
End Property
Public Property Let setIDENTIFICACION2(ByVal dato As String)
        IDENTIFICACION2 = dato
End Property
Public Property Let setINFORMES(ByVal dato As String)
        INFORMES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPC_PEDIDO
'***************************************************************
Public Property Get getID_PEDIDO() As Long
        getID_PEDIDO = ID_PEDIDO
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getTIPO_BOTE_EX_ID() As Long
        getTIPO_BOTE_EX_ID = TIPO_BOTE_EX_ID
End Property
Public Property Get getFECHA_PEDIDO() As String
        getFECHA_PEDIDO = FECHA_PEDIDO
End Property
Public Property Get getIDENTIFICACION() As String
        getIDENTIFICACION = Identificacion
End Property
Public Property Get getCANTIDAD() As Long
        getCANTIDAD = cantidad
End Property
Public Property Get getPEDIDO_CLIENTE() As String
        getPEDIDO_CLIENTE = PEDIDO_CLIENTE
End Property
Public Property Get getFECHA_CLIENTE() As String
        getFECHA_CLIENTE = FECHA_CLIENTE
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getDOC_ID() As Long
        getDOC_ID = DOC_ID
End Property
Public Property Get getNCERTIFICADO() As Long
        getNCERTIFICADO = NCERTIFICADO
End Property
Public Property Get getNORMATIVA_ID() As Long
        getNORMATIVA_ID = NORMATIVA_ID
End Property
Public Property Get getF_CADUCIDAD() As String
        getF_CADUCIDAD = F_CADUCIDAD
End Property
Public Property Get getIDENTIFICACION2() As String
        getIDENTIFICACION2 = IDENTIFICACION2
End Property
Public Property Get getINFORMES() As String
        getINFORMES = INFORMES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPC_PEDIDO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PC_PEDIDO WHERE ID_PEDIDO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PEDIDO = rs("ID_PEDIDO")
                CLIENTE_ID = rs("CLIENTE_ID")
                TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
                FECHA_PEDIDO = rs("FECHA_PEDIDO")
                Identificacion = rs("IDENTIFICACION")
                cantidad = rs("CANTIDAD")
                PEDIDO_CLIENTE = rs("PEDIDO_CLIENTE")
                FECHA_CLIENTE = rs("FECHA_CLIENTE")
                PRECIO = rs("PRECIO")
                PEDIDO_ID = rs("PEDIDO_ID")
                DOC_ID = rs("DOC_ID")
                NCERTIFICADO = rs("NCERTIFICADO")
                NORMATIVA_ID = rs("NORMATIVA_ID")
                If IsNull(rs("F_CADUCIDAD")) Then
                    F_CADUCIDAD = ""
                Else
                    F_CADUCIDAD = rs("F_CADUCIDAD")
                End If
                If IsNull(rs("IDENTIFICACION2")) Then
                    IDENTIFICACION2 = ""
                Else
                    IDENTIFICACION2 = rs("IDENTIFICACION2")
                End If
                If IsNull(rs("INFORMES")) Then
                    INFORMES = ""
                Else
                    INFORMES = rs("INFORMES")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPc_pedido)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PEDIDO) FROM PC_PEDIDO"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PEDIDO = 1
        Else
                ID_PEDIDO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PC_PEDIDO " & _
                   "  VALUES (" & _
                        ID_PEDIDO & "," & CLIENTE_ID & "," & TIPO_BOTE_EX_ID & "," & _
                        "'" & FECHA_PEDIDO & "'" & "," & "'" & Identificacion & "'" & "," & cantidad & "," & _
                        "'" & PEDIDO_CLIENTE & "'" & "," & "'" & FECHA_CLIENTE & "'," & _
                        "'" & PRECIO & "'," & PEDIDO_ID & "," & DOC_ID & "," & NCERTIFICADO & "," & NORMATIVA_ID & "," & _
                        F_CADUCIDAD & ",'" & IDENTIFICACION2 & "','" & INFORMES & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_PEDIDO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPc_pedido)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PC_PEDIDO SET " & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " TIPO_BOTE_EX_ID = " & TIPO_BOTE_EX_ID & "," & _
                        " FECHA_PEDIDO = '" & FECHA_PEDIDO & "'," & _
                        " IDENTIFICACION = '" & Identificacion & "'," & _
                        " CANTIDAD = " & cantidad & "," & _
                        " PEDIDO_CLIENTE = '" & PEDIDO_CLIENTE & "'," & _
                        " FECHA_CLIENTE = '" & FECHA_CLIENTE & "'" & _
                        ",PRECIO = '" & PRECIO & "'" & _
                        ",PEDIDO_ID = " & PEDIDO_ID & _
                        ",NCERTIFICADO = " & NCERTIFICADO & _
                        ",NORMATIVA_ID = " & NORMATIVA_ID & _
                        ",F_CADUCIDAD = " & F_CADUCIDAD & _
                        ",IDENTIFICACION2 = '" & IDENTIFICACION2 & "'" & _
                        ",INFORMES = '" & INFORMES & "'" & _
                " WHERE ID_PEDIDO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPc_pedido)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Me.Carga ID
        If Me.getDOC_ID > 1 Then
            MsgBox "El lote esta insertado en facturación. No se puede eliminar.", vbCritical, App.Title
            Exit Function
        End If
        Dim oPCR As New clsPc_rex
        oPCR.Eliminar ID
        Set oPCR = Nothing
        execute_bd "DELETE FROM PC_PEDIDO WHERE ID_PEDIDO = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPc_pedido)", vbCritical, Err.Description
End Function
Public Function Listado(producto As String, cliente As String, fdesde As String, fhasta As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = ""
        If producto <> "" And producto <> "0" Then
            filtro = filtro & " AND PC.TIPO_BOTE_EX_ID = " & producto
        End If
        If cliente <> "" And cliente <> "0" Then
            filtro = filtro & " AND PC.CLIENTE_ID = " & cliente
        End If
        If fdesde <> "" Then
            filtro = filtro & " AND PC.FECHA_PEDIDO >= '" & Format(fdesde, "yyyy-mm-dd") & "'"
        End If
        If fhasta <> "" Then
            filtro = filtro & " AND PC.FECHA_PEDIDO <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
        End If
        consulta = "SELECT PC.ID_PEDIDO,PC.FECHA_PEDIDO,C.NOMBRE, PC.CANTIDAD,TR.NOMBRE,PC.IDENTIFICACION,PC.PRECIO,CONCAT(CP.CODIGO,' ',CP.DESCRIPCION),PC.PEDIDO_ID " & _
                   "  FROM PC_PEDIDO PC " & _
                   " INNER JOIN CLIENTES C ON PC.CLIENTE_ID = C.ID_CLIENTE " & _
                   " INNER JOIN TIPOS_BOTE_EX TB ON PC.TIPO_BOTE_EX_ID = TB.ID_TIPO_BOTE_EX " & _
                   " INNER JOIN TIPOS_REACTIVO_EX TR ON TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX" & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON PC.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE 1 = 1 " & _
                   filtro & _
                   " ORDER BY PC.FECHA_PEDIDO DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_ID(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT PC.ID_PEDIDO,PC.FECHA_PEDIDO,C.NOMBRE, PC.CANTIDAD,TR.NOMBRE,PC.IDENTIFICACION,PC.PRECIO,CONCAT(CP.CODIGO,' ',CP.DESCRIPCION),PC.PEDIDO_ID " & _
                   "  FROM PC_PEDIDO PC " & _
                   " INNER JOIN CLIENTES C ON PC.CLIENTE_ID = C.ID_CLIENTE " & _
                   " INNER JOIN TIPOS_BOTE_EX TB ON PC.TIPO_BOTE_EX_ID = TB.ID_TIPO_BOTE_EX " & _
                   " INNER JOIN TIPOS_REACTIVO_EX TR ON TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX" & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON PC.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE PC.ID_PEDIDO = " & ID & _
                   " ORDER BY PC.FECHA_PEDIDO DESC"
        Set Listado_por_ID = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PEDIDO,CLIENTE_ID FROM PC_PEDIDO ORDER BY CLIENTE_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function ImprimirEtiquetas(pedido As Long) As Boolean
    Dim prnPrinter As Printer
    
    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros

    log ("Impresion etiqueta PC")
    Dim PC As String
    PC = USUARIO.getUSO
    log ("PC : " & PC)
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_PEQUENA, PC) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Function
    End If
    log ("Parametro impresora : " & Replace(oParametro.getVALOR, "/", "\"))
    
    Dim impresora_encontrada As String
    Dim IMPRESORA As Printer
    impresora_encontrada = ""
    For Each prnPrinter In Printers
'        log "Impresora instalada : " & prnPrinter.DeviceName
        ' Quitamos del nombre de la impresora el #00X para las impresoras TSN del remoto
        Dim impresoraArray() As String
        Dim impresoraNombre As String
        impresoraArray = Split(UCase(prnPrinter.DeviceName), "#")
        impresoraNombre = Trim(impresoraArray(0))
        
'        If UCase(prnPrinter.DeviceName) = Replace(UCase(oParametro.getVALOR), "/", "\") Then
        If impresoraNombre = Replace(UCase(oParametro.getVALOR), "/", "\") Then
            Set Printer = prnPrinter
            Set IMPRESORA = prnPrinter
            impresora_encontrada = prnPrinter.DeviceName
            Exit For
        End If
    Next
    If impresora_encontrada <> "" Then
        Dim imp_ant As String
        imp_ant = Impresora_Predeterminada
        Establecer_Impresora impresora_encontrada
        With frmReport
            .iniciar
            .informe = "\PC\rptPC_ETIQUETA"
            .criterio = "{pc_pedido.ID_PEDIDO} = " & pedido
            .imprimir = True
            .generar
        End With
'        Unload frmReport
        Set frmReport = Nothing
        Establecer_Impresora imp_ant
    Else
        MsgBox "No se localiza la impresora de etiquetas.", vbExclamation, App.Title
    End If
End Function
Public Function ImprimirAlbaran(pedido As Long) As Boolean
    With frmReport
        .iniciar
        .informe = "\PC\rptPC_ALBARAN"
        .criterio = "{pc_pedido.ID_PEDIDO} = " & pedido
        .imprimir = False
        .pdf = ""
        .generar
        .Show 1
'        .Visible = True
    End With
        Unload frmReport
    Set frmReport = Nothing
End Function
Public Function ImprimirCertificado(pedido As Long) As Boolean
    With frmReport
        .iniciar
        .informe = "\PC\rptPC_CERTIFICADO"
        .criterio = "{pc_pedido.ID_PEDIDO} = " & pedido
        .imprimir = False
        .pdf = ""
        .generar
        .Show 1
'        .Visible = True
    End With
        Unload frmReport
    Set frmReport = Nothing
End Function
Public Function Listado_Pendientes_Facturar(CLIENTE_ID As Long, fdesde As String, fhasta As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = ""
        If CLIENTE_ID <> 0 Then
            filtro = filtro & " AND PC.CLIENTE_ID = " & CLIENTE_ID
        End If
        If fdesde <> "" Then
            filtro = filtro & " AND PC.FECHA_PEDIDO >= '" & Format(fdesde, "yyyy-mm-dd") & "'"
        End If
        If fhasta <> "" Then
            filtro = filtro & " AND PC.FECHA_PEDIDO <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
        End If
        consulta = "SELECT C.NOMBRE, TR.NOMBRE,PC.IDENTIFICACION,PC.FECHA_PEDIDO,PC.ID_PEDIDO,PC.CLIENTE_ID,PC.CANTIDAD,PC.PRECIO,CONCAT(CP.CODIGO,' ',CP.DESCRIPCION),PC.PEDIDO_ID " & _
                   "  FROM PC_PEDIDO PC " & _
                   " INNER JOIN CLIENTES C ON PC.CLIENTE_ID = C.ID_CLIENTE " & _
                   " INNER JOIN TIPOS_BOTE_EX TB ON PC.TIPO_BOTE_EX_ID = TB.ID_TIPO_BOTE_EX " & _
                   " INNER JOIN TIPOS_REACTIVO_EX TR ON TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX" & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON PC.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE PC.DOC_ID = 0 " & _
                   filtro & _
                   " ORDER BY PC.ID_PEDIDO"
        Set Listado_Pendientes_Facturar = datos_bd(consulta)
End Function
Public Function Listado_Pendientes_Facturar_ID(ID_PEDIDO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT C.NOMBRE, TR.NOMBRE,PC.IDENTIFICACION,PC.FECHA_PEDIDO,PC.ID_PEDIDO,PC.CLIENTE_ID,PC.CANTIDAD,PC.PRECIO,CONCAT(CP.CODIGO,' ',CP.DESCRIPCION),PC.PEDIDO_ID " & _
                   "  FROM PC_PEDIDO PC " & _
                   " INNER JOIN CLIENTES C ON PC.CLIENTE_ID = C.ID_CLIENTE " & _
                   " INNER JOIN TIPOS_BOTE_EX TB ON PC.TIPO_BOTE_EX_ID = TB.ID_TIPO_BOTE_EX " & _
                   " INNER JOIN TIPOS_REACTIVO_EX TR ON TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX" & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON PC.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE PC.ID_PEDIDO = " & ID_PEDIDO & _
                   " ORDER BY PC.ID_PEDIDO"
        Set Listado_Pendientes_Facturar_ID = datos_bd(consulta)
End Function
Public Function Listado_Para_Factura(pedidos As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT PC.CLIENTE_ID,PC.PEDIDO_ID,PC.FECHA_PEDIDO,CONCAT('SUMINISTRO DE ',TR.NOMBRE),PC.CANTIDAD,PC.PRECIO,PC.CANTIDAD * PC.PRECIO,PC.ID_PEDIDO " & _
                   "  FROM PC_PEDIDO PC " & _
                   " INNER JOIN CLIENTES C ON PC.CLIENTE_ID = C.ID_CLIENTE " & _
                   " INNER JOIN TIPOS_BOTE_EX TB ON PC.TIPO_BOTE_EX_ID = TB.ID_TIPO_BOTE_EX " & _
                   " INNER JOIN TIPOS_REACTIVO_EX TR ON TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX" & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON PC.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE PC.ID_PEDIDO in (" & pedidos & ")" & _
                   " ORDER BY PC.CLIENTE_ID,PC.PEDIDO_ID"
        Set Listado_Para_Factura = datos_bd(consulta)
End Function

Public Function doc_pago(ID_PEDIDO As Long, DOC As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PC_PEDIDO " & _
                   "   SET DOC_ID = " & DOC & _
                   " WHERE ID_PEDIDO = " & ID_PEDIDO & _
                   "   AND DOC_ID = 0"
        execute_bd consulta
        doc_pago = True
        Exit Function
fallo:
        doc_pago = False
        MsgBox "Error al informar el documento de pago (doc_pago)", vbCritical, Err.Description
End Function

Public Function maxNCERTIFICADO(fecha As Date) As Integer
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(NCERTIFICADO) FROM PC_PEDIDO where YEAR(FECHA_PEDIDO) = " & Year(fecha)
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                maxNCERTIFICADO = 1
        Else
                maxNCERTIFICADO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

