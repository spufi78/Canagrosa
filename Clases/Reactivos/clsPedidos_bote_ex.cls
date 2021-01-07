VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedidos_bote_ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PRIORIDAD : 0.BAJA,1.NORMAL,2.URGENTE
' RECIBIDO : 0.PENDIENTE,1.RECIBIDO,2.TRAMITADO
Option Explicit
Private ID_PEDIDO_BOTE_EX As Integer
Private TIPO_BOTE_EX_ID As Integer
Private CODIGO_PEDIDO_PROVEEDOR As Long
Private PROVEEDOR_ID As Integer
Private CENTRO_ID As Integer
Private fecha As String
Private FECHA_PEDIDO As String
Private USUARIO As Integer
Private CANTIDAD As Integer
Private CANTIDAD_RECIBIDA As Integer
Private PRECIO As String
Private dto As String
Private ENVIADO As String
Private CONFIRMADO As String
'M1058-I
Private moneda As Long
'M1058-F
Private anulado As Integer
Private PRIORIDAD As Integer
Private RECIBIDO As Integer
Private FAMILIA_ID As Integer
Private MOTIVO As String
Private OBSERVACIONES As String
Private TRAMITADO_USUARIO_ID As Long
Private TRAMITADO_FECHA As String
'MANTIS-922-I
Private RECEPCION_USUARIO As Long
Private RECEPCION_FECHA As String
'MANTIS-922-F


Public Property Let setID_PEDIDO_BOTE_EX(ByVal dato As Integer)

    ID_PEDIDO_BOTE_EX = dato
End Property
Public Property Let setTIPO_BOTE_EX_ID(ByVal dato As Integer)
    TIPO_BOTE_EX_ID = dato
End Property
Public Property Let setCODIGO_PEDIDO_PROVEEDOR(ByVal dato As Long)
    CODIGO_PEDIDO_PROVEEDOR = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Integer)
    PROVEEDOR_ID = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Integer)
    CENTRO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = dato
End Property
Public Property Let setFECHA_PEDIDO(ByVal dato As String)
    FECHA_PEDIDO = dato
End Property
Public Property Let setUSUARIO(ByVal dato As Integer)
    USUARIO = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Integer)
    CANTIDAD = dato
End Property
Public Property Let setCANTIDAD_RECIBIDA(ByVal dato As Integer)
    CANTIDAD_RECIBIDA = dato
End Property
Public Property Let setPRIORIDAD(ByVal dato As Integer)
    PRIORIDAD = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = Trim(dato)
End Property
Public Property Let setRECIBIDO(ByVal dato As Integer)
    RECIBIDO = Trim(dato)
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Integer)
    FAMILIA_ID = Trim(dato)
End Property
Public Property Let setMOTIVO(ByVal dato As String)
    MOTIVO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = dato
End Property
Public Property Let setTRAMITADO_USUARIO_ID(ByVal dato As Long)
    TRAMITADO_USUARIO_ID = dato
End Property
Public Property Let setTRAMITADO_FECHA(ByVal dato As String)
    TRAMITADO_FECHA = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = dato
End Property
Public Property Let setDTO(ByVal dato As String)
    dto = Trim(dato)
End Property
Public Property Let setENVIADO(ByVal dato As String)
    ENVIADO = Trim(dato)
End Property
Public Property Let setCONFIRMADO(ByVal dato As String)
    CONFIRMADO = Trim(dato)
End Property
'M1058-I
Public Property Let setMONEDA(ByVal dato As Long)
    moneda = Trim(dato)
End Property
'M1058-F
'MANTIS-922-I
Public Property Let setRECEPCION_USUARIO(ByVal dato As Long)
    RECEPCION_USUARIO = dato
End Property
Public Property Let setRECEPCION_FECHA(ByVal dato As String)
    RECEPCION_FECHA = dato
End Property

'MANTIS-922-F



Public Property Get getID_PEDIDO_BOTE_EX() As Integer
    getID_PEDIDO_BOTE_EX = ID_PEDIDO_BOTE_EX
End Property
Public Property Get getTIPO_BOTE_EX_ID() As Integer
    getTIPO_BOTE_EX_ID = TIPO_BOTE_EX_ID
End Property
Public Property Get getCODIGO_PEDIDO_PROVEEDOR() As Long
    getCODIGO_PEDIDO_PROVEEDOR = CODIGO_PEDIDO_PROVEEDOR
End Property
Public Property Get getPROVEEDOR_ID() As Integer
    getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getCENTRO_ID() As Integer
    getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getFECHA_PEDIDO() As String
    getFECHA_PEDIDO = FECHA_PEDIDO
End Property
Public Property Get getUSUARIO() As Integer
    getUSUARIO = USUARIO
End Property
Public Property Get getCANTIDAD() As Integer
    getCANTIDAD = CANTIDAD
End Property
Public Property Get getCANTIDAD_RECIBIDA() As Integer
    getCANTIDAD_RECIBIDA = CANTIDAD_RECIBIDA
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Property Get getPRIORIDAD() As Integer
    getPRIORIDAD = PRIORIDAD
End Property
Public Property Get getRECIBIDO() As Integer
    getRECIBIDO = RECIBIDO
End Property
Public Property Get getFAMILIA_ID() As Integer
    getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getMOTIVO() As String
    getMOTIVO = MOTIVO
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getTRAMITADO_USUARIO_ID() As Long
    getTRAMITADO_USUARIO_ID = TRAMITADO_USUARIO_ID
End Property
Public Property Get getTRAMITADO_FECHA() As String
    getTRAMITADO_FECHA = TRAMITADO_FECHA
End Property
Public Property Get getPRECIO() As String
    getPRECIO = PRECIO
End Property
Public Property Get getDTO() As String
    getDTO = dto
End Property
Public Property Get getENVIADO() As String
    getENVIADO = ENVIADO
End Property
Public Property Get getCONFIRMADO() As String
    getCONFIRMADO = CONFIRMADO
End Property
'M1058-I
Public Property Get getMONEDA() As Long
    getMONEDA = moneda
End Property
'M1058-F
'MANTIS-922-I
Public Property Get getRECEPCION_USUARIO() As Long
    getRECEPCION_USUARIO = RECEPCION_USUARIO
End Property
Public Property Get getRECEPCION_FECHA() As String
    getRECEPCION_FECHA = RECEPCION_FECHA
End Property
'MANTIS-922-F

Public Function Listado(proveedor As Long, TIPO_REACTIVO As Long, fdesde As Date, fhasta As Date, ESTADO As Integer, pedido As String, CENTRO_ID As Integer) As ADODB.Recordset
    'M1076 : NºPedido proveedor
    Dim consulta As String
    Dim filtro As String
    If proveedor <> 0 Then
        filtro = filtro & " AND PB.PROVEEDOR_ID = " & proveedor
    End If
    If TIPO_REACTIVO <> 0 Then
        filtro = filtro & " AND TB.TIPO_REACTIVO_EX_ID = " & TIPO_REACTIVO
    End If
    'M1076-I
    If pedido <> "" Then
        filtro = filtro & " AND PB.CODIGO_PEDIDO_PROVEEDOR = " & pedido
    End If
    'M1076-F
    filtro = filtro & " AND PB.FECHA >= ' " & Format(fdesde, "yyyy-mm-dd") & "'"
    filtro = filtro & " AND PB.FECHA <= ' " & Format(fhasta, "yyyy-mm-dd") & "'"
    If ESTADO >= 0 Then
        filtro = filtro & " AND PB.RECIBIDO = " & ESTADO
    End If
    If CENTRO_ID <> 0 Then
        filtro = filtro & " AND PB.CENTRO_ID = " & CENTRO_ID
    End If
    
    consulta = "SELECT PB.ID_PEDIDO_BOTE_EX,P.NOMBRE,TB.CODIGO,TR.NOMBRE, PB.FECHA,PB.CANTIDAD,PB.TIPO_BOTE_EX_ID,EM.USUARIO,PB.RECIBIDO,PB.PRIORIDAD,PB.CANTIDAD_RECIBIDA,PB.CODIGO_PEDIDO_PROVEEDOR,C.NOMBRE,PB.MOTIVO,FAM.NOMBRE,TB.TIPO_M_REFERENCIA_ID " & _
               "  FROM PEDIDOS_BOTE_EX PB " & _
               " INNER JOIN PROVEEDORES P ON PB.PROVEEDOR_ID = P.ID_PROVEEDOR " & _
               " INNER JOIN TIPOS_BOTE_EX TB ON PB.TIPO_BOTE_EX_ID = TB.ID_TIPO_BOTE_EX " & _
               " INNER JOIN TIPOS_REACTIVO_EX TR ON TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX " & _
               " INNER JOIN usuarios EM ON PB.USUARIO = EM.ID_EMPLEADO " & _
               " INNER JOIN CENTROS C ON PB.CENTRO_ID = C.ID_CENTRO " & _
               "  LEFT JOIN FAMILIAS FAM ON PB.FAMILIA_ID = FAM.ID_FAMILIA " & _
               " WHERE PB.ANULADO = 0 " & _
               filtro & _
               " order by PB.ID_PEDIDO_BOTE_EX DESC"
    Set Listado = datos_bd(consulta)
End Function
Public Function calcularID_PEDIDO(ID_PEDIDO As Long, TIPO_BOTE_EX_ID As Long) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim ID As Long
   On Error GoTo calcularID_PEDIDO_Error

    ID = 0
    consulta = "SELECT * FROM PEDIDOS_BOTE_EX " & _
               " WHERE ID_PEDIDO_BOTE_EX=" & ID_PEDIDO & _
               "   AND TIPO_BOTE_EX_ID = " & TIPO_BOTE_EX_ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        consulta = "SELECT min(id_pedido_bote_ex) from pedidos_bote_ex A " & _
                   " WHERE A.CODIGO_PEDIDO_PROVEEDOR=" & rs("CODIGO_PEDIDO_PROVEEDOR") & _
                   "   AND YEAR(A.FECHA_PEDIDO) =" & Year(rs("FECHA_PEDIDO"))
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            ID = rs(0)
        End If
    End If
    calcularID_PEDIDO = ID

   On Error GoTo 0
   Exit Function

calcularID_PEDIDO_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularID_PEDIDO of Módulo de clase clsPedidos_bote_ex"
End Function
Public Function CARGAR(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM PEDIDOS_BOTE_EX " & _
               " WHERE ID_PEDIDO_BOTE_EX=" & ID & " ORDER BY RECIBIDO DESC"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_PEDIDO_BOTE_EX = rs("ID_PEDIDO_BOTE_EX")
        TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
        CODIGO_PEDIDO_PROVEEDOR = rs("CODIGO_PEDIDO_PROVEEDOR")
        PROVEEDOR_ID = rs("PROVEEDOR_ID")
        CENTRO_ID = rs("CENTRO_ID")
        fecha = rs("FECHA")
        If Not IsNull(rs("fecha_pedido")) Then
            FECHA_PEDIDO = rs("FECHA_PEDIDO")
        End If
        USUARIO = rs("USUARIO")
        CANTIDAD = rs("CANTIDAD")
        CANTIDAD_RECIBIDA = rs("CANTIDAD_RECIBIDA")
        anulado = rs("ANULADO")
        PRIORIDAD = rs("PRIORIDAD")
        RECIBIDO = rs("RECIBIDO")
        FAMILIA_ID = rs("FAMILIA_ID")
        MOTIVO = rs("MOTIVO")
        OBSERVACIONES = rs("OBSERVACIONES")
        TRAMITADO_USUARIO_ID = rs("TRAMITADO_USUARIO_ID")
        If IsNull(rs("TRAMITADO_FECHA")) Then
            TRAMITADO_FECHA = ""
        Else
            TRAMITADO_FECHA = rs("TRAMITADO_FECHA")
        End If
        RECEPCION_USUARIO = rs("RECEPCION_USUARIO")
        If IsNull(rs("RECEPCION_FECHA")) Then
            RECEPCION_FECHA = ""
        Else
            RECEPCION_FECHA = rs("RECEPCION_FECHA")
        End If

        PRECIO = rs("PRECIO")
        dto = rs("DTO")
        ENVIADO = rs("ENVIADO")
        CONFIRMADO = rs("CONFIRMADO")
        'M1058-I
        moneda = rs("MONEDA")
        'M1058-F
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar el pedido de reactivo.", vbCritical, Err.Description
End Function
Public Function cargar_con_bote(ID As Long, BOTE As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM PEDIDOS_BOTE_EX " & _
               " WHERE ID_PEDIDO_BOTE_EX=" & ID & _
               "   AND TIPO_BOTE_EX_ID=" & BOTE
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_PEDIDO_BOTE_EX = rs("ID_PEDIDO_BOTE_EX")
        TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
        CODIGO_PEDIDO_PROVEEDOR = rs("CODIGO_PEDIDO_PROVEEDOR")
        PROVEEDOR_ID = rs("PROVEEDOR_ID")
        CENTRO_ID = rs("CENTRO_ID")
        fecha = rs("FECHA")
        FECHA_PEDIDO = rs("FECHA_PEDIDO")
        USUARIO = rs("USUARIO")
        CANTIDAD = rs("CANTIDAD")
        CANTIDAD_RECIBIDA = rs("CANTIDAD_RECIBIDA")
        anulado = rs("ANULADO")
        PRIORIDAD = rs("PRIORIDAD")
        RECIBIDO = rs("RECIBIDO")
        FAMILIA_ID = rs("FAMILIA_ID")
        MOTIVO = rs("MOTIVO")
        OBSERVACIONES = rs("OBSERVACIONES")
        TRAMITADO_USUARIO_ID = rs("TRAMITADO_USUARIO_ID")
        If IsNull(rs("TRAMITADO_FECHA")) Then
            TRAMITADO_FECHA = ""
        Else
            TRAMITADO_FECHA = rs("TRAMITADO_FECHA")
        End If
        PRECIO = rs("PRECIO")
        dto = rs("DTO")
        ENVIADO = rs("ENVIADO")
        CONFIRMADO = rs("CONFIRMADO")
        'M1058-I
        moneda = rs("MONEDA")
        'M1058-F
        cargar_con_bote = True
    Else
        cargar_con_bote = False
    End If
    Exit Function
fallo:
    cargar_con_bote = False
    MsgBox "Error al cargar el pedido de reactivo.", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PEDIDO_BOTE_EX) FROM PEDIDOS_BOTE_EX"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PEDIDO_BOTE_EX = 1
    Else
        ID_PEDIDO_BOTE_EX = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    If ID_PEDIDO_BOTE_EX = 0 Then
        Me.CrearID
    End If
    If FECHA_PEDIDO = "" Then
        Me.setFECHA_PEDIDO = "0000-00-00"
    End If
'MANTIS-922-I
'    consulta = "INSERT INTO PEDIDOS_BOTE_EX " & _
'               " VALUES(" & _
'               ID_PEDIDO_BOTE_EX & "," & _
'               TIPO_BOTE_EX_ID & "," & _
'               CODIGO_PEDIDO_PROVEEDOR & "," & _
'               PROVEEDOR_ID & ",'" & _
'               fecha & "','" & _
'               FECHA_PEDIDO & "'," & _
'               USUARIO & "," & _
'               cantidad & "," & _
'               CANTIDAD_RECIBIDA & ",'" & _
'               PRECIO & "','" & _
'               dto & "'," & _
'               ANULADO & "," & _
'               PRIORIDAD & "," & _
'              RECIBIDO & "," & _
'               "'" & MOTIVO & "'," & _
'               "'" & OBSERVACIONES & "','" & _
'              ENVIADO & "','" & _
'               CONFIRMADO & "'," & _
'               "0,CURRENT_TIMESTAMP" & ")"
'M1058 : CAMPO MONEDA
    consulta = "INSERT INTO PEDIDOS_BOTE_EX " & _
               " VALUES(" & _
               ID_PEDIDO_BOTE_EX & "," & _
               TIPO_BOTE_EX_ID & "," & _
               CODIGO_PEDIDO_PROVEEDOR & "," & _
               PROVEEDOR_ID & "," & _
               CENTRO_ID & ",'" & _
               fecha & "','" & _
               FECHA_PEDIDO & "'," & _
               USUARIO & "," & _
               CANTIDAD & "," & _
               CANTIDAD_RECIBIDA & ",'" & _
               PRECIO & "','" & _
               dto & "'," & _
               anulado & "," & _
               PRIORIDAD & "," & _
               RECIBIDO & "," & _
               FAMILIA_ID & "," & _
               "'" & MOTIVO & "'," & _
               "'" & OBSERVACIONES & "','" & _
               ENVIADO & "','" & _
               CONFIRMADO & "'," & _
               moneda & "," & _
               "0,CURRENT_TIMESTAMP,0,null" & ")"
'MANTIS-922-F
    execute_bd consulta
    Insertar = ID_PEDIDO_BOTE_EX
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Pedido (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM PEDIDOS_BOTE_EX WHERE ID_PEDIDO_BOTE_EX=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el Bote de Reactivo (Eliminar)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, BOTE As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    'M1005 INFORMAR EL PRECIO Y LA CANTIDAD
    'M1058 MONEDA
'                   "      ,MOTIVO     = '" & MOTIVO & "' "

    consulta = "UPDATE PEDIDOS_BOTE_EX " & _
               "   SET PROVEEDOR_ID = " & PROVEEDOR_ID & _
               "      ,CENTRO_ID = " & CENTRO_ID & _
               "      ,FAMILIA_ID = " & FAMILIA_ID & _
               "      ,DTO        = '" & moneda_bd(dto) & "' " & _
               "      ,FECHA_PEDIDO = '" & FECHA_PEDIDO & "' " & _
               "      ,ENVIADO    = '" & ENVIADO & "' " & _
               "      ,CONFIRMADO = '" & CONFIRMADO & "' " & _
               "      ,MONEDA = " & moneda & _
               "      ,PRECIO = '" & PRECIO & "' " & _
               "      ,CANTIDAD = " & CANTIDAD & _
               "      ,OBSERVACIONES = '" & OBSERVACIONES & "' " & _
               " WHERE ID_PEDIDO_BOTE_EX = " & ID & _
               "   AND TIPO_BOTE_EX_ID = " & BOTE
               
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el Pedido (Recibir)", vbCritical, Err.Description
End Function
Public Function recibir_bote(ID As Long, BOTE As Long) As Boolean
'MANTIS-922: Añadidas las columnas RECEPCION_FECHA y RECEPCION_USUARIO a la consulta de actualización
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE PEDIDOS_BOTE_EX " & _
               "   SET CANTIDAD_RECIBIDA = CANTIDAD_RECIBIDA + 1 " & _
               "   ,RECEPCION_FECHA =  current_timestamp " & _
               "   ,RECEPCION_USUARIO = " & globales.USUARIO.getID_EMPLEADO & _
               " WHERE ID_PEDIDO_BOTE_EX = " & ID & _
               "   AND TIPO_BOTE_EX_ID = " & BOTE
    execute_bd consulta
    recibir_bote = True
    Exit Function
fallo:
    recibir_bote = False
    MsgBox "Error al recibir_bote Del Pedido (recibir_bote)", vbCritical, Err.Description
End Function

Public Function Recibir(ID As Long, BOTE As Long, EVALUAR As Boolean) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    'JGM : FALTA LA FECHA DE RECEPCION Y EL USUARIO
    consulta = "UPDATE PEDIDOS_BOTE_EX " & _
               "   SET RECIBIDO = 1 " & _
               "     ,CANTIDAD_RECIBIDA = CANTIDAD " & _
               "     ,RECEPCION_FECHA =  current_timestamp " & _
               "     ,RECEPCION_USUARIO = " & globales.USUARIO.getID_EMPLEADO & _
               " WHERE ID_PEDIDO_BOTE_EX = " & ID & _
               "   AND TIPO_BOTE_EX_ID = " & BOTE
    execute_bd consulta
    ' Enviar al responsable si lo tiene
    Dim oTb As New clsTipos_bote_ex
    oTb.CARGAR BOTE
    If oTb.getRESPONSABLE_ID <> 0 And EVALUAR = True Then
        Dim oUsuario As New clsUsuarios
        Dim oTR As New clsTipos_reactivo_ex
        Dim ASUNTO As String
        Dim texto As String
        Dim mailto As String
        oUsuario.CARGAR oTb.getRESPONSABLE_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        ASUNTO = "Evaluación de reactivo : " & oTR.getNOMBRE
        texto = "Hay que realizar la evaluación del reactivo del asunto."
        If oUsuario.getEMAIL <> "" Then
            mailto = oUsuario.getEMAIL & ";"
        End If
        Dim oParametro As New clsParametros
        oParametro.Carga parametros.REX_PEDIDO_DESTINATARIO_CORREO, ""
        If oParametro.getVALOR <> "" Then
            mailto = mailto & oParametro.getVALOR
        End If
        Call Enviar_Mail_CDO(mailto, ASUNTO, texto, vbNullString)
'        genera_correo mailto, asunto, texto, "", 0
    End If
    Recibir = True
    Exit Function
fallo:
    Recibir = False
    MsgBox "Error al Recibir el Pedido (Recibir)", vbCritical, Err.Description
End Function
Public Function Anular(pedido As Long, BOTE As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE PEDIDOS_BOTE_EX SET " & _
               " ANULADO = 1" & _
               " WHERE ID_PEDIDO_BOTE_EX = " & pedido & _
               "   AND TIPO_BOTE_EX_ID = " & BOTE
    execute_bd consulta
    Anular = True
    Exit Function
fallo:
    Anular = False
    MsgBox "Error al anular el Pedido (Anular)", vbCritical, Err.Description
End Function
Public Function Tramitar(pedido As Long, BOTE As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE PEDIDOS_BOTE_EX " & _
               "   SET RECIBIDO = 2 " & _
               "     , TRAMITADO_USUARIO_ID = " & globales.USUARIO.getID_EMPLEADO & _
               "     , TRAMITADO_FECHA = CURRENT_TIMESTAMP " & _
               " WHERE ID_PEDIDO_BOTE_EX = " & pedido & _
               "   AND TIPO_BOTE_EX_ID = " & BOTE
    execute_bd consulta
    Tramitar = True
    Exit Function
fallo:
    Tramitar = False
    MsgBox "Error al Tramitar el Pedido (Tramitar)", vbCritical, Err.Description
End Function
Public Function informar_pedido_proveedor(pedido As Long, BOTE As Long, CODIGO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE PEDIDOS_BOTE_EX " & _
               "   SET CODIGO_PEDIDO_PROVEEDOR = " & CODIGO & _
               "      ,FECHA_PEDIDO = CURRENT_DATE " & _
               " WHERE ID_PEDIDO_BOTE_EX = " & pedido & _
               "   AND TIPO_BOTE_EX_ID = " & BOTE & _
               "   AND CODIGO_PEDIDO_PROVEEDOR = 0"
    execute_bd consulta
    informar_pedido_proveedor = True
    Exit Function
fallo:
    informar_pedido_proveedor = False
    MsgBox "Error al informar_pedido_proveedor (Tramitar)", vbCritical, Err.Description
End Function

Public Function proveedor_pedido(pedido As Long, ANNO As Integer) As Long
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT PROVEEDOR_ID FROM PEDIDOS_BOTE_EX " & _
               " WHERE CODIGO_PEDIDO_PROVEEDOR=" & pedido & _
               " and year(FECHA_PEDIDO) = " & ANNO
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        proveedor_pedido = rs(0)
    Else
        proveedor_pedido = 0
    End If
    Exit Function
fallo:
    proveedor_pedido = 0
    MsgBox "Error al cargar el pedido de reactivo.", vbCritical, Err.Description
End Function


Public Function Numero_Pedido() As Integer
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(CODIGO_PEDIDO_PROVEEDOR) " & _
               "  FROM PEDIDOS_BOTE_EX " & _
               " WHERE CODIGO_PEDIDO_PROVEEDOR <> 99999 and YEAR(fecha_pedido) = " & Year(Date)
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Numero_Pedido = 1
    Else
        Numero_Pedido = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Function
Public Function imprimir(pedido As Long, ANNO As Integer, pdf As String) As Boolean
    With frmReport
        .iniciar
        .informe = "\REX\rptREX_Pedido"
        ' M1058 : Moneda
        .criterio = "{pedidos_bote_ex.CODIGO_PEDIDO_PROVEEDOR} = " & pedido & " and year({pedidos_bote_ex.FECHA_PEDIDO}) = " & ANNO & " and {pedidos_bote_ex.ANULADO} = 0.00 and {decodificadora.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_MONEDA
        .imprimir = False
        .pdf = pdf
        .generar
        If pdf = "" Then
            .Show 1
            Unload frmReport
        End If
    End With
End Function
Public Function CARGAR_POR_PEDIDO_PROVEEDOR(CODIGO As Long, ANNO As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT C.NOMBRE AS REACTIVO,D.NOMBRE AS FAMILIA, B.CODIGO_PROVEEDOR,A.* " & _
               "  FROM PEDIDOS_BOTE_EX A " & _
               " INNER JOIN TIPOS_BOTE_EX B ON  A.TIPO_BOTE_EX_ID = B.ID_TIPO_BOTE_EX " & _
               " INNER JOIN TIPOS_REACTIVO_EX C ON B.TIPO_REACTIVO_EX_ID = C.ID_TIPO_REACTIVO_EX " & _
               "  LEFT JOIN FAMILIAS D ON A.FAMILIA_ID = D.ID_FAMILIA " & _
               " WHERE A.CODIGO_PEDIDO_PROVEEDOR=" & CODIGO & _
               "   AND YEAR(A.FECHA_PEDIDO) =" & ANNO & _
               "   AND A.ANULADO = 0 " & _
               " ORDER BY A.id_pedido_bote_ex ASC"
    Set CARGAR_POR_PEDIDO_PROVEEDOR = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar el pedido de reactivo.", vbCritical, Err.Description
End Function
Public Function ListadoPedidosRealizados(proveedor As Long, fecha_i As String, fecha_f As String, MOTIVO As String, CENTRO_ID As String) As ADODB.Recordset
On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim strProveedor As String
    Dim srtMotivo As String
    If proveedor <> 0 Then
        strProveedor = " AND C.ID_PROVEEDOR = " & proveedor
    Else
        strProveedor = " "
    End If
    If MOTIVO <> "" Then
        srtMotivo = srtMotivo & " AND a.MOTIVO LIKE '%" & MOTIVO & "%'"
    End If
    If CENTRO_ID <> "" And CENTRO_ID <> "0" Then
        srtMotivo = srtMotivo & " AND a.CENTRO_ID = " & CENTRO_ID
    End If
    
    fecha_i = Format(fecha_i, "yyyy-mm-dd")
    fecha_f = Format(fecha_f, "yyyy-mm-dd")
    ' Solicitado por Jeni 01/03/2018
    ' Incluir los pedidos recibidos o no, incluir el centro y el número de pedido
    consulta = "SELECT A.ID_PEDIDO_BOTE_EX, B.ID_TIPO_BOTE_EX, ct.NOMBRE,a.CODIGO_PEDIDO_PROVEEDOR,a.FECHA,a.FECHA_PEDIDO,c.NOMBRE, d.NOMBRE,b.CODIGO_PROVEEDOR, a.MOTIVO, b.CANTIDAD, sum(a.CANTIDAD), a.PRECIO, sum(a.CANTIDAD * a.PRECIO)" & _
               " FROM pedidos_bote_ex a" & _
               " LEFT JOIN tipos_bote_ex b on a.TIPO_BOTE_EX_ID = b.id_tipo_bote_ex" & _
               " LEFT JOIN tipos_reactivo_ex d on b.TIPO_REACTIVO_EX_ID = d.ID_TIPO_REACTIVO_EX" & _
               " LEFT JOIN proveedores c on a.PROVEEDOR_ID = c.ID_PROVEEDOR" & _
               " LEFT JOIN centros ct ON a.CENTRO_ID = ct.ID_CENTRO " & _
               " WHERE a.TRAMITADO_FECHA >= '" & fecha_i & "' and a.TRAMITADO_FECHA <= '" & fecha_f & "'" & _
               " AND a.RECIBIDO <> 0" & _
               " AND a.ANULADO = 0" & _
               strProveedor & srtMotivo & _
               " GROUP by ct.nombre,a.CODIGO_PEDIDO_PROVEEDOR,a.FECHA,a.FECHA_PEDIDO,c.NOMBRE, d.NOMBRE, a.MOTIVO, b.CANTIDAD " & _
               " order by ct.NOMBRE,a.CODIGO_PEDIDO_PROVEEDOR "
'               " AND a.RECIBIDO = 1"
    Set ListadoPedidosRealizados = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar el listado de pedidos", vbCritical, Err.Description
End Function
Public Function ListadoPedidosFechaPrevista(proveedor As Long, fecha_i As String, fecha_f As String, CENTRO_ID As String) As ADODB.Recordset
On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim strProveedor As String
    Dim srtMotivo As String
    If proveedor <> 0 Then
        strProveedor = " AND b.ID_PROVEEDOR = " & proveedor
    Else
        strProveedor = " "
    End If
    If CENTRO_ID <> "" And CENTRO_ID <> "0" Then
        srtMotivo = srtMotivo & " AND a.CENTRO_ID = " & CENTRO_ID
    End If
    
    fecha_i = Format(fecha_i, "yyyy-mm-dd")
    fecha_f = Format(fecha_f, "yyyy-mm-dd")
    
'               "      ,concat(substr(a.CONFIRMADO,7,4),'-',substr(a.CONFIRMADO,4,2),'-',substr(a.CONFIRMADO,1,2)) AS FECHA_PREVISTA "
    consulta = "SELECT a.ID_PEDIDO_BOTE_EX,ct.nombre,b.NOMBRE, d.NOMBRE, a.CODIGO_PEDIDO_PROVEEDOR,a.FECHA_PEDIDO " & _
               "      ,a.CONFIRMADO, b.DIAS_ENTREGA, date_add(a.FECHA_PEDIDO,interval b.DIAS_ENTREGA DAY) AS FECHA_ESTIMADA,IFNULL(a.RECEPCION_FECHA,'') " & _
               "      , if(a.CONFIRMADO <> '', " & _
               "          (if(concat(substr(a.CONFIRMADO,7,4),'-',substr(a.CONFIRMADO,4,2),'-',substr(a.CONFIRMADO,1,2)) < current_date,1,0)), " & _
               "          (if(date_add(a.FECHA_PEDIDO,interval b.DIAS_ENTREGA DAY) < current_date,1,0)) " & _
               "        ) AS CADUCADO " & _
               " FROM pedidos_bote_ex a, proveedores b, tipos_bote_ex c, tipos_reactivo_ex d,centros ct " & _
               " WHERE a.FECHA_PEDIDO >= '" & fecha_i & "' and a.FECHA_PEDIDO <= '" & fecha_f & "'" & _
               " AND a.RECIBIDO <> 0 AND a.ANULADO = 0 " & _
               " AND a.PROVEEDOR_ID = b.ID_PROVEEDOR " & _
               " and a.TIPO_BOTE_EX_ID = c.ID_TIPO_BOTE_EX " & _
               " and c.TIPO_REACTIVO_EX_ID = d.ID_TIPO_REACTIVO_EX " & _
               " and a.centro_id = ct.id_centro " & _
               strProveedor & srtMotivo & _
               " ORDER BY a.FECHA_PEDIDO "
    Set ListadoPedidosFechaPrevista = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar el listado de pedidos por Fecha Prevista", vbCritical, Err.Description
End Function

Public Function ListadoRelaciones(proveedor As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT CONCAT(LPAD(PB.CODIGO_PEDIDO_PROVEEDOR,4,'0'),'-',YEAR(PB.FECHA_PEDIDO)),PB.FECHA_PEDIDO " & _
               "  FROM PEDIDOS_BOTE_EX PB " & _
               " WHERE PB.ANULADO = 0 AND PB.CODIGO_PEDIDO_PROVEEDOR <> 0 " & _
               "   AND PB.PROVEEDOR_ID =  " & proveedor & _
               " order by PB.ID_PEDIDO_BOTE_EX DESC"
    Set ListadoRelaciones = datos_bd(consulta)
End Function


