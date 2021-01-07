VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBotes_ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_BOTE_EX As Long
Private NUMERO As Integer
Private CODIGO As String
Private TIPO_BOTE_EX_ID As Integer
Private LOTE As String
Private FECHA_PEDIDO As String
Private FECHA_RECEPCION As String
Private FECHA_APERTURA As String
Private fecha_caducidad As String
Private NO_CADUCA As Integer
Private NO_CONFORME As Integer
Private HENKEL As Integer
Private fecha_fin As String
Private anulado As Integer
Private ABIERTO As Integer
Private FINALIZADO As Integer
Private CERTIFICADO_EXTERNO As String
Private OBSERVACIONES As String
'M1076-I
Private PEDIDO_BOTE_EX_ID As Long
Private CENTRO_ID As Long
'M1076-F
'M1166-I
Private CONVERTIDO As Integer
Private CERTIFICADO As Integer
Private USUARIO_CERTIFICADOR As Integer
Private FECHA_CERTIFICACION As String
'M1166-F
Public Sub Imprimir_Listado(ByVal prm_criterio As String)
    Dim objRpt As New frmReport
    Dim arrNom() As String, arrVal() As String
    
    ReDim arrNom(1)
    ReDim arrVal(1)
    
    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
    
    With objRpt
        .iniciar
        .informe = "REX\rptListadoBotesReactivosExternos"
        .criterio = prm_criterio
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal
        
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
            
        Unload objRpt
        Set objRpt = Nothing
        
End Sub

Public Property Let setID_BOTE_EX(ByVal dato As Long)
    ID_BOTE_EX = dato
End Property
Public Property Let setNUMERO(ByVal dato As Integer)
    NUMERO = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
    CODIGO = dato
End Property
Public Property Let setTIPO_BOTE_EX_ID(ByVal dato As Integer)
    TIPO_BOTE_EX_ID = dato
End Property
Public Property Let setLOTE(ByVal dato As String)
    LOTE = dato
End Property
Public Property Let setFECHA_PEDIDO(ByVal dato As String)
    FECHA_PEDIDO = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
    FECHA_RECEPCION = dato
End Property
Public Property Let setFECHA_APERTURA(ByVal dato As String)
    FECHA_APERTURA = Trim(dato)
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
    fecha_caducidad = Trim(dato)
End Property
Public Property Let setNO_CADUCA(ByVal dato As Integer)
    NO_CADUCA = Trim(dato)
End Property
Public Property Let setNO_CONFORME(ByVal dato As Integer)
    NO_CONFORME = Trim(dato)
End Property
Public Property Let setHENKEL(ByVal dato As Integer)
    HENKEL = Trim(dato)
End Property
Public Property Let setFECHA_FIN(ByVal dato As String)
    fecha_fin = Trim(dato)
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
Public Property Let setABIERTO(ByVal dato As Integer)
    ABIERTO = dato
End Property
Public Property Let setFINALIZADO(ByVal dato As Integer)
    FINALIZADO = dato
End Property
Public Property Let setCERTIFICADO_EXTERNO(ByVal dato As String)
    CERTIFICADO_EXTERNO = Replace(Trim(dato), "\", "/")
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
'M1076-I
Public Property Let setPEDIDO_BOTE_EX_ID(ByVal dato As Long)
    PEDIDO_BOTE_EX_ID = Trim(dato)
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
    CENTRO_ID = Trim(dato)
End Property
'M1076-F
'M1166-I
Public Property Let setCONVERTIDO(ByVal dato As Integer)
    CONVERTIDO = dato
End Property
Public Property Let setCERTIFICADO(ByVal dato As Integer)
    CERTIFICADO = dato
End Property
Public Property Let setUSUARIO_CERTIFICADOR(ByVal dato As Integer)
    USUARIO_CERTIFICADOR = dato
End Property
Public Property Let setFECHA_CERTIFICACION(ByVal dato As String)
    FECHA_CERTIFICACION = Trim(dato)
End Property
'M1166-F

Public Property Get getID_BOTE_EX() As Long
    getID_BOTE_EX = ID_BOTE_EX
End Property
Public Property Get getNUMERO() As Integer
    getNUMERO = NUMERO
End Property
Public Property Get getCODIGO() As String
    getCODIGO = CODIGO
End Property
Public Property Get getTIPO_BOTE_EX_ID() As Integer
    getTIPO_BOTE_EX_ID = TIPO_BOTE_EX_ID
End Property
Public Property Get getLOTE() As String
    getLOTE = LOTE
End Property
Public Property Get getFECHA_PEDIDO() As String
    getFECHA_PEDIDO = FECHA_PEDIDO
End Property
Public Property Get getFECHA_RECEPCION() As String
    getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getFECHA_APERTURA() As String
    getFECHA_APERTURA = FECHA_APERTURA
End Property
Public Property Get getFECHA_CADUCIDAD() As String
    getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getNO_CADUCA() As Integer
    getNO_CADUCA = NO_CADUCA
End Property
Public Property Get getNO_CONFORME() As Integer
    getNO_CONFORME = NO_CONFORME
End Property
Public Property Get getHENKEL() As Integer
    getHENKEL = HENKEL
End Property
Public Property Get getFECHA_FIN() As String
    getFECHA_FIN = fecha_fin
End Property
Public Property Get getANULADO() As String
    getANULADO = anulado
End Property
Public Property Get getABIERTO() As Integer
    getABIERTO = ABIERTO
End Property
Public Property Get getFINALIZADO() As Integer
    getFINALIZADO = FINALIZADO
End Property
Public Property Get getCERTIFICADO_EXTERNO() As String
    getCERTIFICADO_EXTERNO = Replace(CERTIFICADO_EXTERNO, "/", "\")
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
'M1076-I
Public Property Get getPEDIDO_BOTE_EX_ID() As Long
    getPEDIDO_BOTE_EX_ID = PEDIDO_BOTE_EX_ID
End Property
Public Property Get getCENTRO_ID() As Long
    getCENTRO_ID = CENTRO_ID
End Property
'M1076-F
'M1166-I
Public Property Get getCONVERTIDO() As Integer
    getCONVERTIDO = CONVERTIDO
End Property
Public Property Get getCERTIFICADO() As Integer
    getCERTIFICADO = CERTIFICADO
End Property
Public Property Get getUSUARIO_CERTIFICADOR() As Integer
    getUSUARIO_CERTIFICADOR = USUARIO_CERTIFICADOR
End Property
Public Property Get getFECHA_CERTIFICACION() As String
    getFECHA_CERTIFICACION = FECHA_CERTIFICACION
End Property
'M1166-F

Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM BOTES_EX order by ID_BOTE_EX ASC"
    Set Listado = datos_bd(consulta)
End Function
Public Function CARGAR(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM BOTES_EX WHERE ID_BOTE_EX=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        ID_BOTE_EX = rs("ID_BOTE_EX")
        NUMERO = rs("NUMERO")
        CODIGO = rs("CODIGO")
        TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
        LOTE = rs("LOTE")
        If Not IsNull(rs("fecha_pedido")) Then
            FECHA_PEDIDO = rs("FECHA_PEDIDO")
        End If
        If Not IsNull(rs("fecha_recepcion")) Then
            FECHA_RECEPCION = rs("FECHA_RECEPCION")
        End If
        If Not IsNull(rs("fecha_apertura")) Then
            FECHA_APERTURA = rs("FECHA_APERTURA")
        End If
        If Not IsNull(rs("fecha_caducidad")) Then
            fecha_caducidad = rs("FECHA_CADUCIDAD")
        End If
        If Not IsNull(rs("fecha_fin")) Then
            fecha_fin = rs("FECHA_FIN")
        End If
        NO_CADUCA = rs("NO_CADUCA")
        NO_CONFORME = rs("NO_CONFORME")
        HENKEL = rs("HENKEL")
        anulado = rs("ANULADO")
        FINALIZADO = rs("FINALIZADO")
        ABIERTO = rs("ABIERTO")
        CERTIFICADO_EXTERNO = rs("CERTIFICADO_EXTERNO")
        OBSERVACIONES = rs("OBSERVACIONES")
        'M1076-I
        PEDIDO_BOTE_EX_ID = rs("PEDIDO_BOTE_EX_ID")
        CENTRO_ID = rs("CENTRO_ID")
        'M1076-F
        'M1166-I
        CONVERTIDO = rs("CONVERTIDO")
        CERTIFICADO = rs("CERTIFICADO")
        USUARIO_CERTIFICADOR = rs("USUARIO_CERTIFICADOR")
        If Not IsNull(rs("FECHA_CERTIFICACION")) Then
            FECHA_CERTIFICACION = rs("FECHA_CERTIFICACION")
        End If
        'M1166-F
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar el bote de reactivo.", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_BOTE_EX) FROM BOTES_EX"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_BOTE_EX = 1
    Else
        ID_BOTE_EX = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    If ID_BOTE_EX = 0 Then
        Me.CrearID
    End If
    ' Crear numero particular
    CalcularNumero (TIPO_BOTE_EX_ID)
' J51
    Me.setFECHA_FIN = "0000-00-00"
    ' M1076 : PEDIDO_BOTE_EX_ID
    ' M1166 : CONVERTIDO Y CERTIFICADO
    '         VALORES OBLIGATORIOS EN LA CREACIÓN DEL BOTE
    '         USUARIO CERTIFICADOR SE OBTIENE DEL TIPO
    usuarioCertificador
    'M1166-F
    consulta = "Insert into BOTES_EX " & _
               " values(" & _
               ID_BOTE_EX & "," & _
               NUMERO & ",'" & _
               CODIGO & "'," & _
               TIPO_BOTE_EX_ID & ",'" & _
               LOTE & "','" & _
               FECHA_PEDIDO & "','" & _
               FECHA_RECEPCION & "','" & _
               FECHA_APERTURA & "','" & _
               fecha_caducidad & "'," & _
               NO_CADUCA & "," & _
               NO_CONFORME & "," & _
               HENKEL & ",'" & _
               fecha_fin & "'," & anulado & "," & _
               ABIERTO & "," & FINALIZADO & ",'" & _
               CERTIFICADO_EXTERNO & "','" & _
               OBSERVACIONES & "'," & _
               PEDIDO_BOTE_EX_ID & "," & _
               CENTRO_ID & "," & _
               CONVERTIDO & "," & _
               CERTIFICADO & "," & _
               USUARIO_CERTIFICADOR & ",'" & _
               FECHA_CERTIFICACION & "')"
    execute_bd consulta
    Insertar = ID_BOTE_EX
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Reactivo (Insertar)", vbCritical, Err.Description
End Function
'M1166-I
Private Sub usuarioCertificador()
    Dim oTIPO As New clsTipos_bote_ex
    oTIPO.CARGAR CLng(TIPO_BOTE_EX_ID)
    USUARIO_CERTIFICADOR = oTIPO.getRESPONSABLE_ID
    CONVERTIDO = 1
    CERTIFICADO = 0
End Sub
'M1166-F
Public Function CalcularNumero(TIPO_BOTE As Long) As Integer
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim oTb As New clsTipos_bote_ex
    oTb.CARGAR TIPO_BOTE
    Dim oDeco As New clsDecodificadora
    oDeco.Carga_valor DECODIFICADORA.REX_TIPOS, oTb.getTIPO_M_REFERENCIA_ID
    consulta = "SELECT MAX(NUMERO) " & _
               "  FROM BOTES_EX " & _
               " WHERE CODIGO='" & oDeco.getPARAMETROS & "'" & _
               "   AND YEAR(FECHA_RECEPCION) = " & Year(Date)
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        NUMERO = 1
    Else
        NUMERO = rs.Fields(0) + 1
    End If
    CODIGO = oDeco.getPARAMETROS
    Set rs = Nothing
End Function

Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM BOTES_EX WHERE ID_BOTE_EX=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el Bote de Reactivo (Eliminar)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX SET " & _
               "  LOTE = '" & LOTE & "'," & _
               "  CENTRO_ID = " & CENTRO_ID & "," & _
               "  FECHA_RECEPCION = '" & FECHA_RECEPCION & "'," & _
               "  FECHA_APERTURA = '" & FECHA_APERTURA & "'," & _
               "  FECHA_CADUCIDAD = '" & fecha_caducidad & "'," & _
               "  NO_CADUCA = " & NO_CADUCA & "," & _
               "  NO_CONFORME = " & NO_CONFORME & "," & _
               "  HENKEL = " & HENKEL & "," & _
               "  FECHA_FIN = '" & fecha_fin & "'," & _
               "  ANULADO = " & anulado & "," & _
               "  ABIERTO = " & ABIERTO & "," & _
               "  FINALIZADO = " & FINALIZADO & "," & _
               "  OBSERVACIONES = '" & OBSERVACIONES & "'" & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar el Reactivo (Modificar)", vbCritical, Err.Description
End Function
Public Function InformarRutaCertificado(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX SET " & _
               "  CERTIFICADO_EXTERNO = '" & CERTIFICADO_EXTERNO & "'" & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    InformarRutaCertificado = True
    Exit Function
fallo:
    InformarRutaCertificado = False
    MsgBox "Error al Informar la ruta del certificado externo (Modificar)", vbCritical, Err.Description
End Function

Public Function Abrir(ID As Long, fecha As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX SET " & _
               " ABIERTO = 1, FECHA_APERTURA = '" & Format(fecha, "yyyy-mm-dd") & "'" & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    Abrir = True
    
    Dim oRA As New clsReactivos_almacen
    oRA.dar_apertura_bote_salida_reactivo ID, "R", USUARIO.getID_EMPLEADO
    Set oRA = Nothing
    
    Exit Function
fallo:
    Abrir = False
    MsgBox "Error al Abrir el Bote (Abrir)", vbCritical, Err.Description
End Function
Public Function No_Conforme_Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX SET " & _
               " NO_CONFORME = 1 " & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    No_Conforme_Modificar = True
    
    Exit Function
fallo:
    No_Conforme_Modificar = False
    MsgBox "Error al No_Conforme_Modificar el Bote (Abrir)", vbCritical, Err.Description
End Function
Public Function observacion(ID As Long, texto As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX SET " & _
               " OBSERVACIONES = concat(OBSERVACIONES, '" & texto & "')" & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    observacion = True
    
    Exit Function
fallo:
    observacion = False
    MsgBox "Error al observacion el Bote (observacion)", vbCritical, Err.Description
End Function

Public Function AnularTerminacion(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX SET " & _
               " FECHA_FIN = '0000-00-00',FINALIZADO=0" & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    AnularTerminacion = True
    
    Dim oRA As New clsReactivos_almacen
    
    oRA.anular_dar_finalizacion_reactivo ID, "R", USUARIO.getID_EMPLEADO
    Set oRA = Nothing
    
    Exit Function
fallo:
    AnularTerminacion = False
    MsgBox "Error al AnularTerminacion el Bote (AnularTerminacion)", vbCritical, Err.Description
End Function

Public Function Terminar(ID As Long, fecha As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX SET " & _
               " FECHA_FIN = '" & Format(fecha, "yyyy-mm-dd") & "',FINALIZADO=1" & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    Terminar = True
    
    Dim oRA As New clsReactivos_almacen
    
    oRA.dar_finalizacion_reactivo ID, "R", USUARIO.getID_EMPLEADO
    Set oRA = Nothing
    
    Exit Function
fallo:
    Terminar = False
    MsgBox "Error al Terminar el Bote (Terminar)", vbCritical, Err.Description
End Function
Public Function Listado_Para_Propio(id_reactivo_ex As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT be.id_bote_ex, " & _
               "       pro.nombre, " & _
               "       tb.codigo, " & _
               "       be.LOTE, " & _
               "       be.fecha_apertura, " & _
               "       be.fecha_fin, " & _
               "       be.fecha_caducidad, " & _
               "       be.tipo_bote_ex_id, " & _
               "       tb.precio " & _
               " FROM BOTES_EX be, " & _
               "      TIPOS_BOTE_EX tb, " & _
               "      proveedores pro, " & _
               "      TIPOS_REACTIVO_EX tr " & _
               " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
               "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
               "   AND tr.ID_TIPO_REACTIVO_EX = " & id_reactivo_ex & _
               "   and tb.proveedor_id = pro.id_proveedor " & _
               "   AND be.ANULADO = 0 " & _
               "   AND be.FINALIZADO = 0 " & _
               "   AND be.ABIERTO = 1 " & _
               " ORDER BY be.id_bote_ex desc"
    Set Listado_Para_Propio = datos_bd(consulta)
End Function
Public Function imprimir_etiqueta(LISTA_REACTIVOS As String, Optional EQUIPO As String) As Boolean
    Dim prnPrinter As Printer
    
    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros

   On Error GoTo imprimir_etiqueta_Error
    log ("Impresion etiqueta reactivo-I")
    Dim PC As String
    If EQUIPO <> "" Then
        PC = EQUIPO
    Else
        PC = USUARIO.getUSO
    End If
    Dim i As Integer
    If PC = "IBERIA" Then
        Dim oEtiqueta As New clsEtiquetas
        Dim Listar() As String
        Listar = Split(LISTA_REACTIVOS, ",")
        For i = LBound(Listar) To UBound(Listar)
            With oEtiqueta
                .setTIPO_ID = C_ETIQUETAS_TIPOS.C_ETIQUETAS_TIPOS_REX
                .setCENTRO_ID = CENTROS.CENTRO_MADRID
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setID = Listar(i)
                .Insertar
            End With
        Next
        MsgBox "Se han enviado a imprimir las copias.", vbInformation, App.Title
        Exit Function
    End If
    
'    log ("PC : " & PC)
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_MEDIANA, PC) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Function
    End If
'    Firmas.copiar_firma_usuario_activo
'    log ("Parametro impresora : " & Replace(oParametro.getVALOR, "/", "\"))
    
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
        log ("Impresion etiqueta reactivo-impresora_encontrada")
        Dim imp_ant As String
        Dim P1() As String
        Dim P2() As String
        ReDim P1(1) As String
        ReDim P2(1) As String
        P1(1) = "FIRMA"
        If ReadINI(App.Path + "\config.ini", "Documentos", "firmas") = "c:\geslab\recursos\firmas" Then
            Dim firma As String
            Dim lista() As String
            If USUARIO.getFIRMA <> "" Then
                lista = Split(USUARIO.getFIRMA, "/")
                firma = ReadINI(App.Path + "\config.ini", "Documentos", "firmas") & "\" & lista(UBound(lista))
            End If
            P2(1) = firma
        Else
            P2(1) = Replace(USUARIO.getFIRMA, "/", "\")
        End If
        imp_ant = Impresora_Predeterminada
        Establecer_Impresora impresora_encontrada
        log ("Impresion etiqueta reactivo-impresora_encontrada-I")
        With frmReport
            .iniciar
            .informe = "\REX\rptREX_Etiqueta_Pequeña"
'            .informe = "\REX\rptREX_Etiqueta"
            .criterio = "{botes_ex.ID_BOTE_EX} in [" & LISTA_REACTIVOS & "]"
            .ParametrosNombre = P1
            .ParametrosValores = P2
            .imprimir = True
'            .imprimir = False
            .generar
'            .Show 1
        End With
        Unload frmReport
        Set frmReport = Nothing
        Establecer_Impresora imp_ant
        log ("Impresion etiqueta reactivo-impresora_encontrada-F")
    Else
        MsgBox "No se localiza la impresora de etiquetas.", vbExclamation, App.Title
    End If
    log ("Impresion etiqueta reactivo-F")

   On Error GoTo 0
   Exit Function

imprimir_etiqueta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Imprimir_Etiqueta of Módulo de clase clsBotes_ex"
End Function

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim consulta As String
    consulta = "SELECT A.ID_BOTE_EX, CONCAT(C.NOMBRE, '  (',CONCAT('Num.',CAST(A.ID_BOTE_EX AS CHAR)), CONCAT('  Lote:',A.LOTE),') -> Centro : ', D.NOMBRE) " & _
               "  FROM BOTES_EX A,TIPOS_BOTE_EX B, TIPOS_REACTIVO_EX C, CENTROS D " & _
               " Where A.TIPO_BOTE_EX_ID = B.ID_TIPO_BOTE_EX " & _
               "   AND B.TIPO_REACTIVO_EX_ID = C.ID_TIPO_REACTIVO_EX " & _
               "   AND A.CENTRO_ID = D.ID_CENTRO " & _
               "   AND A.ANULADO = 0 " & _
               filtro
'               "   AND A.ANULADO = 0 AND FECHA_FIN ='0000-00-00' "
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "BOTES_EX"
        .setDESCRIPCION = "Reactivos"
        .setPK = "ID_EQUIPO_EX"
        .setCAMPO = "CONCAT(C.NOMBRE, '  (',CONCAT('Num.',CAST(A.ID_BOTE_EX AS CHAR)), CONCAT('  Lote:',A.LOTE),')')"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Certificar(ID As Long, CONFORME As Boolean, fecha As Date) As Boolean
    On Error GoTo fallo
    Dim consulta As String
     
    consulta = "UPDATE BOTES_EX SET " & _
               " CERTIFICADO = " & IIf(CONFORME = True, "1", "2") & _
               " ,FECHA_CERTIFICACION = '" & Format(fecha, "yyyy-mm-dd") & "'" & _
               " ,USUARIO_CERTIFICADOR = " & USUARIO.getID_EMPLEADO & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    Certificar = True
    
    Exit Function
fallo:
    Certificar = False
    MsgBox "Error al Certificar el Bote", vbCritical, Err.Description
End Function
Public Function ReCertificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
     
    consulta = "UPDATE BOTES_EX SET " & _
               " CERTIFICADO = 0" & _
               " ,FECHA_CERTIFICACION = '0000-00-00'" & _
               " WHERE ID_BOTE_EX = " & ID
    execute_bd consulta
    ReCertificar = True
    
    Exit Function
fallo:
    ReCertificar = False
    MsgBox "Error al ReCertificar el Bote", vbCritical, Err.Description
End Function

Public Function listadoExistenciasPorTipo(TIPO_BOTE As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT be.id_bote_ex, " & _
               "       tb.codigo, " & _
               "       tr.nombre, " & _
               "       be.fecha_recepcion, " & _
               "       be.fecha_apertura, " & _
               "       be.fecha_fin, " & _
               "       be.fecha_caducidad, " & _
               "       be.tipo_bote_ex_id, " & _
               "       be.LOTE, " & _
               "       pbe.precio / tb.cantidad_unidad_pedido, " & _
               "       tb.cantidad, tb.tipo_m_referencia_id, be.numero,be.codigo,be.no_conforme,be.NO_CADUCA " & _
               " FROM BOTES_EX be, " & _
               "      TIPOS_BOTE_EX tb, " & _
               "      TIPOS_REACTIVO_EX tr, " & _
               "      PEDIDOS_BOTE_EX pbe " & _
               " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
               "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
               "   AND be.PEDIDO_BOTE_EX_ID = pbe.ID_PEDIDO_BOTE_EX and be.tipo_bote_ex_id = pbe.tipo_bote_ex_id " & _
               "   AND be.anulado = 0 " & _
               "   AND be.finalizado = 0 " & _
               "   AND be.tipo_bote_ex_id = " & TIPO_BOTE & _
               " ORDER BY be.id_bote_ex desc"
    Set listadoExistenciasPorTipo = datos_bd(consulta)
End Function

