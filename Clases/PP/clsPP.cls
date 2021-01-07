VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPP
'***************************************************************
Private ID_PP As Long
Private NUMERO As Integer
Private ANNO As Integer
Private PRESUPUESTO As String
Private OBSERVACIONES As String
Private PROVEEDOR_ID As Long
Private CENTRO_ID As Long
Private fecha_creacion As String
Private FECHA_ENVIO As String
Private FECHA_RECEPCION As String
Private FACTURA_RECIBIDA As Integer
Private NFACTURA As String
Private ffactura As String
Private USUARIO_ID As Long
Private ESTADO As Long
Private tipo As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPP
'***************************************************************
Public Property Let setID_PP(ByVal dato As Long)
        ID_PP = dato
End Property
Public Property Let setNUMERO(ByVal dato As Integer)
        NUMERO = dato
End Property
Public Property Let setANNO(ByVal dato As Integer)
        ANNO = dato
End Property
Public Property Let setPRESUPUESTO(ByVal dato As String)
        PRESUPUESTO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
        CENTRO_ID = dato
End Property
Public Property Let setFECHA_CREACION(ByVal dato As String)
        fecha_creacion = dato
End Property
Public Property Let setFECHA_ENVIO(ByVal dato As String)
        FECHA_ENVIO = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        FECHA_RECEPCION = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setFACTURA_RECIBIDA(ByVal dato As Long)
        FACTURA_RECIBIDA = dato
End Property
Public Property Let setNFACTURA(ByVal dato As String)
        NFACTURA = dato
End Property
Public Property Let setFFACTURA(ByVal dato As String)
        ffactura = dato
End Property
Public Property Let setESTADO(ByVal dato As Long)
        ESTADO = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPP
'***************************************************************
Public Property Get getID_PP() As Long
        getID_PP = ID_PP
End Property
Public Property Get getNUMERO() As Integer
        getNUMERO = NUMERO
End Property
Public Property Get getANNO() As Integer
        getANNO = ANNO
End Property
Public Property Get getPRESUPUESTO() As String
        getPRESUPUESTO = PRESUPUESTO
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getCENTRO_ID() As Long
        getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getFECHA_CREACION() As String
        getFECHA_CREACION = fecha_creacion
End Property
Public Property Get getFECHA_ENVIO() As String
        getFECHA_ENVIO = FECHA_ENVIO
End Property
Public Property Get getFECHA_RECEPCION() As String
        getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFACTURA_RECIBIDA() As Long
        getFACTURA_RECIBIDA = FACTURA_RECIBIDA
End Property
Public Property Get getNFACTURA() As String
        getNFACTURA = NFACTURA
End Property
Public Property Get getFFACTURA() As String
        getFFACTURA = ffactura
End Property
Public Property Get getESTADO() As Long
        getESTADO = ESTADO
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPP
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * FROM PP WHERE ID_PP = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_PP = rs("ID_PP")
            NUMERO = rs("NUMERO")
            ANNO = rs("ANNO")
            PRESUPUESTO = rs("PRESUPUESTO")
            OBSERVACIONES = rs("OBSERVACIONES")
            PROVEEDOR_ID = rs("PROVEEDOR_ID")
            CENTRO_ID = rs("CENTRO_ID")
            fecha_creacion = Format(rs("FECHA_CREACION"), "yyyy-mm-dd")
            FECHA_ENVIO = Format(rs("FECHA_ENVIO"), "yyyy-mm-dd")
            FECHA_RECEPCION = Format(rs("FECHA_RECEPCION"), "yyyy-mm-dd")
            FACTURA_RECIBIDA = rs("FACTURA_RECIBIDA")
            NFACTURA = rs("NFACTURA")
            ffactura = Format(rs("FFACTURA"), "yyyy-mm-dd")
            USUARIO_ID = rs("USUARIO_ID")
            ESTADO = rs("ESTADO")
            tipo = rs("TIPO")
            
            Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPP)", vbCritical, Err.Description
End Function

Public Function CrearID()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT MAX(ID_PP) FROM PP"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            ID_PP = 1
    Else
            ID_PP = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Function

Public Function CrearNumero() As String
    Dim rs As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT MAX(NUMERO) FROM PP WHERE ANNO = " & Year(fecha_creacion)
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            NUMERO = 1
    Else
            NUMERO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_PP = 0 Then
            Me.CrearID
            Me.CrearNumero
        End If
        consulta = "INSERT INTO PP (" & _
                    "      ID_PP, NUMERO, ANNO, PRESUPUESTO " & _
                    "    , OBSERVACIONES, PROVEEDOR_ID,CENTRO_ID " & _
                    "    , FECHA_CREACION, FECHA_ENVIO, FECHA_RECEPCION " & _
                    "    , FACTURA_RECIBIDA, NFACTURA, FFACTURA, USUARIO_ID, ESTADO, TIPO" & _
                    "    , FTIMESTP " & _
                    ") VALUES (" & _
                           ID_PP & "," & NUMERO & "," & Year(fecha_creacion) & "," & _
                           "'" & PRESUPUESTO & "'," & "'" & OBSERVACIONES & "'" & ", " & _
                           PROVEEDOR_ID & "," & CENTRO_ID & "," & "CURRENT_TIMESTAMP," & FECHA_ENVIO & "," & FECHA_RECEPCION & "," & _
                           FACTURA_RECIBIDA & ", '" & NFACTURA & "'," & "'" & ffactura & "'," & usuario.getID_EMPLEADO & "," & _
                           ESTADO & "," & tipo & "," & _
                           "CURRENT_TIMESTAMP)"
        execute_bd consulta
        Insertar = ID_PP
        Exit Function
        
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPP)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE PP SET " & _
                   "       PRESUPUESTO = '" & PRESUPUESTO & "'" & _
                   "     , PROVEEDOR_ID = " & PROVEEDOR_ID & _
                   "     , CENTRO_ID = " & CENTRO_ID & _
                   "     , OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                   "     , NFACTURA = '" & NFACTURA & "'" & _
                   "     , FFACTURA = '" & ffactura & "'" & _
                   "     , FACTURA_RECIBIDA = " & FACTURA_RECIBIDA & _
                   "     , FECHA_ENVIO = " & FECHA_ENVIO & _
                   "     , FECHA_RECEPCION = " & FECHA_RECEPCION & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE ID_PP = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPP)", vbCritical, Err.Description
End Function
Public Function Tramitar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE PP SET " & _
                   "       ESTADO = 1" & _
                   "     , APROBADOR_ID = " & usuario.getID_EMPLEADO & _
                   "     , FECHA_APROBACION = '" & Format(Date, "yyyy-mm-dd") & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE ID_PP = " & ID
        
        execute_bd consulta
        Tramitar = True
        Exit Function
fallo:
        Tramitar = False
        MsgBox "Error al Tramitar (clsPP)", vbCritical, Err.Description
End Function
Public Function Enviar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim oPP As New clsPP
        If oPP.Carga(ID) Then
            If oPP.getFECHA_ENVIO <> "" Then
                MsgBox "ERROR : El pedido ya esta marcado como Enviado.", vbCritical, App.Title
                Enviar = False
                Exit Function
            Else
                consulta = "UPDATE PP SET " & _
                           "       FECHA_ENVIO = '" & Format(Date, "yyyy-mm-dd") & "'" & _
                           "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                           " WHERE ID_PP = " & ID
                execute_bd consulta
            End If
        End If
        Enviar = True
        Exit Function
fallo:
        Enviar = False
        MsgBox "Error al Enviar (clsPP)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim oAdjuntos As New clsAdjuntos
        oAdjuntos.EliminarTodos TOBJETO.TOBJETO_PEDIDO_PROVEEDOR, ID
        Set oAdjuntos = Nothing
        ' se quitan las determinaciones del detalle del paquete
        execute_bd "DELETE FROM PP_DETALLE WHERE PP_ID = " & ID
        ' se elimina el paquete
        execute_bd "DELETE FROM PP WHERE ID_PP = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPP)", vbCritical, Err.Description
End Function
Public Function Listado(NUMERO As String, usuario As String, CENTRO_ID As String, PROVEEDOR_ID As String, strFechaDesde As String, strFechaHasta As String, FAMILIA_ID As Long, DESCRIPCION As String) As ADODB.Recordset
    Dim filtro As String
    Dim consulta As String
    filtro = ""
    If NUMERO <> "" Then
        filtro = filtro & " AND PP.NUMERO = " & NUMERO
    End If
    If usuario <> "" Then
        filtro = filtro & " AND U.NOMBRE LIKE '%" & usuario & "%' "
    End If
    If CENTRO_ID <> "" And CENTRO_ID <> "0" Then
        filtro = filtro & " AND PP.CENTRO_ID = " & CENTRO_ID
    End If
    If PROVEEDOR_ID <> "" And PROVEEDOR_ID <> "0" Then
        filtro = filtro & " AND PP.PROVEEDOR_ID = " & PROVEEDOR_ID
    End If
    If FAMILIA_ID <> 0 Then
        filtro = filtro & " AND PPD.FAMILIA_ID = " & FAMILIA_ID
    End If
    If DESCRIPCION <> "" Then
        filtro = filtro & " AND PPD.descripcion = '" & DESCRIPCION & "'"
    End If
    consulta = "SELECT DISTINCT PP.ID_PP,PP.NUMERO,PP.ANNO,PR.NOMBRE,PP.PRESUPUESTO " & _
                   " , PP.FECHA_CREACION, PP.FECHA_ENVIO, PP.FECHA_RECEPCION " & _
                   " , C.NOMBRE , U.NOMBRE, PR.ID_PROVEEDOR " & _
                   "  FROM PP " & _
                   "  LEFT JOIN PROVEEDORES PR ON PP.PROVEEDOR_ID = PR.ID_PROVEEDOR " & _
                   "  LEFT JOIN usuarios U ON PP.USUARIO_ID = U.ID_EMPLEADO " & _
                   "  LEFT JOIN CENTROS C ON PP.CENTRO_ID = C.ID_CENTRO " & _
                   "  LEFT JOIN PP_DETALLE PPD ON PP.ID_PP = PPD.PP_ID " & _
                   " WHERE PP.FECHA_CREACION >= '" & strFechaDesde & "' " & _
                   "   AND PP.FECHA_CREACION <= '" & strFechaHasta & "' " & _
                   filtro & _
                   " ORDER BY PP.ID_PP DESC "
    Set Listado = datos_bd(consulta)
End Function
Public Function ListadoID(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT PP.ID_PP,PP.NUMERO,PP.ANNO,PR.NOMBRE,PP.PRESUPUESTO " & _
                   " , PP.FECHA_CREACION, PP.FECHA_ENVIO, PP.FECHA_RECEPCION " & _
                   " , C.NOMBRE , U.NOMBRE, PR.ID_PROVEEDOR " & _
                   "  FROM PP " & _
                   "  LEFT JOIN PROVEEDORES PR ON PP.PROVEEDOR_ID = PR.ID_PROVEEDOR " & _
                   "  LEFT JOIN usuarios U ON PP.USUARIO_ID = U.ID_EMPLEADO " & _
                   "  LEFT JOIN CENTROS C ON PP.CENTRO_ID = C.ID_CENTRO " & _
                   " WHERE PP.ID_PP = " & ID
    Set ListadoID = datos_bd(consulta)
End Function

Public Function imprimir(PAQUETE As Long, pdf As String) As Boolean
        Dim oPaquete As New clsPP
        oPaquete.Carga PAQUETE
                
        With frmReport
            .iniciar
            .informe = "\Pedidos\rptPedidos_Proveedor"
            .criterio = "{PP.ID_PP}=" & PAQUETE & " and {fp.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & " and {vencimiento.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_VENCIMIENTOS
            .imprimir = False
            If Trim(pdf) <> "" Then
              .pdf = pdf
            End If
            .generar
            .Visible = False
'            .Show 1
        End With
        Unload frmReport
End Function

Public Function ListadoRelaciones(proveedor As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT P.ID_PP,CONCAT(P.NUMERO,'/',P.ANNO),P.FECHA_CREACION " & _
               "  FROM PP P " & _
               " WHERE P.ESTADO = 1 " & _
               "   AND P.PROVEEDOR_ID =  " & proveedor & _
               " order by P.ID_PP DESC"
    Set ListadoRelaciones = datos_bd(consulta)
End Function


