VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEP_Paquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEP_PAQUETES
'***************************************************************
Private ID_PAQUETE As Long
Private ASUNTO As String
Private DETALLE As String
Private tipo As Integer
Private CLIENTE_ID As Long
Private MENSAJERIA_ID As Long
Private fecha_creacion As String
Private HORA_CREACION As String
Private fecha_RECEPCION As String
Private USUARIO_ID As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEP_PAQUETES
'***************************************************************
Public Property Let setID_PAQUETE(ByVal dato As Long)
        ID_PAQUETE = dato
End Property
Public Property Let setASUNTO(ByVal dato As String)
        ASUNTO = dato
End Property
Public Property Let setDETALLE(ByVal dato As String)
        DETALLE = dato
End Property
Public Property Let setTIPO(ByVal dato As Integer)
        tipo = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setMENSAJERIA_ID(ByVal dato As Long)
        MENSAJERIA_ID = dato
End Property
Public Property Let setFECHA_CREACION(ByVal dato As String)
        fecha_creacion = dato
End Property
Public Property Let setHORA_CREACION(ByVal dato As String)
        HORA_CREACION = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        fecha_RECEPCION = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEP_PAQUETES
'***************************************************************
Public Property Get getID_PAQUETE() As Long
        getID_PAQUETE = ID_PAQUETE
End Property
Public Property Get getASUNTO() As String
        getASUNTO = ASUNTO
End Property
Public Property Get getDETALLE() As String
        getDETALLE = DETALLE
End Property
Public Property Get getTIPO() As Integer
        getTIPO = tipo
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getMENSAJERIA_ID() As Long
        getMENSAJERIA_ID = MENSAJERIA_ID
End Property
Public Property Get getFECHA_CREACION() As String
        getFECHA_CREACION = fecha_creacion
End Property
Public Property Get getHORA_CREACION() As String
        getHORA_CREACION = HORA_CREACION
End Property
Public Property Get getFECHA_RECEPCION() As String
        getFECHA_RECEPCION = fecha_RECEPCION
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEP_PAQUETES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * FROM EP_PAQUETES WHERE ID_PAQUETE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PAQUETE = rs("ID_PAQUETE")
                ASUNTO = rs("ASUNTO")
                DETALLE = rs("DETALLE")
                tipo = rs("TIPO")
                CLIENTE_ID = rs("CLIENTE_ID")
                MENSAJERIA_ID = rs("MENSAJERIA_ID")
                fecha_creacion = Format(rs("FECHA_CREACION"), "yyyy-mm-dd")
                HORA_CREACION = rs("HORA_CREACION")
                If Not IsNull(rs("FECHA_RECEPCION")) Then
                    fecha_RECEPCION = rs("FECHA_RECEPCION")
                End If
                USUARIO_ID = rs("USUARIO_ID")

                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEP_Paquetes)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_PAQUETE) FROM EP_PAQUETES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PAQUETE = 1
        Else
                ID_PAQUETE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        If ID_PAQUETE = 0 Then
            Me.CrearID
        End If
        
        consulta = "INSERT INTO EP_PAQUETES (" & _
                   "      ID_PAQUETE, ASUNTO, DETALLE,TIPO " & _
                   "    , CLIENTE_ID, MENSAJERIA_ID " & _
                   "    , FECHA_CREACION, HORA_CREACION, FECHA_RECEPCION, USUARIO_ID " & _
                   "    , CUSERID, FTIMESTP " & _
                   ") VALUES (" & _
                   ID_PAQUETE & "," & "'" & ASUNTO & "'" & ", " & "'" & DETALLE & "', " & tipo & "," & _
                   CLIENTE_ID & ", " & MENSAJERIA_ID & "," & _
                   "'" & Format(fecha_creacion, "yyyy-mm-dd") & "'" & ", " & "'" & HORA_CREACION & "', " & _
                   IIf(fecha_RECEPCION = "", "null", "'" & Format(fecha_RECEPCION, "yyyy-mm-dd") & "'") & "," & _
                   USUARIO_ID & ", " & _
                   "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP)"
        
        execute_bd consulta
        Insertar = ID_PAQUETE
        Exit Function
        
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEP_Paquetes)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EP_PAQUETES " & _
                   "   SET ASUNTO = '" & ASUNTO & "'" & _
                   "     , DETALLE = '" & DETALLE & "'" & _
                   "     , TIPO = " & tipo & _
                   "     , CLIENTE_ID = " & CLIENTE_ID & _
                   "     , MENSAJERIA_ID = " & MENSAJERIA_ID & _
                   "     , FECHA_RECEPCION = " & IIf(fecha_RECEPCION = "", "null", "'" & Format(fecha_RECEPCION, "yyyy-mm-dd") & "'") & _
                   " WHERE ID_PAQUETE = " & ID
                   
'                   "     , USUARIO_ID = " & USUARIO_ID
'                   "     , CUSERID = '" & USUARIO.getUSUARIO & "'"
'                   "     , FTIMESTP = CURRENT_TIMESTAMP "

        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEP_Paquetes)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        ' se elimina el paquete
        execute_bd "DELETE FROM EP_PAQUETES WHERE ID_PAQUETE = " & ID
        Eliminar = True
        Exit Function
        
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEP_Paquetes)", vbCritical, Err.Description
End Function

Public Function Listado(tipo As Integer, strAsunto As String, strCliente As String, strMensajeria As String, strFechaDesde As String, strFechaHasta As String, strUsuario As String) As ADODB.Recordset
    Dim consulta As String
    Dim s As String
    s = ""
    If strAsunto <> "" Then
        s = s & "   AND P.ASUNTO LIKE '%" & strAsunto & "%' "
    End If
    If strUsuario <> "" Then
        s = s & "   AND U.NOMBRE LIKE '%" & strUsuario & "%' "
    End If
    
    If tipo = 0 Then
        consulta = "SELECT P.ID_PAQUETE, CN.ID_CLIENTE, CN.NOMBRE CLIENTE " & _
                   "     , P.ASUNTO, M.DESCRIPCION MENSAJERIA " & _
                   "     , P.FECHA_CREACION, P.HORA_CREACION  " & _
                   "     , U.NOMBRE USUARIO, P.FECHA_RECEPCION " & _
                   "  FROM EP_PAQUETES P, CLIENTES CN, usuarios U, decodificadora M " & _
                   " WHERE P.CLIENTE_ID = CN.ID_CLIENTE " & _
                   "   AND P.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND M.CODIGO = " & DECODIFICADORA.EP_EMPRESAS_MENSAJERIA & _
                   "   AND P.MENSAJERIA_ID = M.VALOR " & _
                   IIf(strMensajeria <> "" And strMensajeria <> "0", "   AND P.MENSAJERIA_ID = " & strMensajeria, "") & _
                   s & _
                   IIf(strCliente <> "0", "   AND CN.ID_CLIENTE = " & strCliente, "") & _
                   "   AND P.FECHA_CREACION >= '" & strFechaDesde & "' " & _
                   "   AND P.FECHA_CREACION <= '" & strFechaHasta & "' " & _
                   "   AND P.TIPO = 0 " & _
                   " ORDER BY P.ID_PAQUETE DESC "
    Else
        consulta = "SELECT P.ID_PAQUETE, CN.ID_PROVEEDOR, CN.NOMBRE CLIENTE " & _
                   "     , P.ASUNTO, M.DESCRIPCION MENSAJERIA " & _
                   "     , P.FECHA_CREACION, P.HORA_CREACION  " & _
                   "     , U.NOMBRE USUARIO, P.FECHA_RECEPCION " & _
                   "  FROM EP_PAQUETES P, PROVEEDORES CN, usuarios U, decodificadora M " & _
                   " WHERE P.CLIENTE_ID = CN.ID_PROVEEDOR " & _
                   "   AND P.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND M.CODIGO = " & DECODIFICADORA.EP_EMPRESAS_MENSAJERIA & _
                   "   AND P.MENSAJERIA_ID = M.VALOR " & _
                   IIf(strMensajeria <> "" And strMensajeria <> "0", "   AND P.MENSAJERIA_ID = " & strMensajeria, "") & _
                   s & _
                   IIf(strCliente <> "0", "   AND CN.ID_PROVEEDOR = " & strCliente, "") & _
                   "   AND P.FECHA_CREACION >= '" & strFechaDesde & "' " & _
                   "   AND P.FECHA_CREACION <= '" & strFechaHasta & "' " & _
                   "   AND P.TIPO = 1 " & _
                   " ORDER BY P.ID_PAQUETE DESC "
    
    End If
    Set Listado = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT ID_PAQUETE, ASUNTO FROM EP_PAQUETES ORDER BY ASUNTO"
        Set Listado_Combo = datos_bd(consulta)
        
End Function

Public Function enviar_correo(lngPaquete As Long) As Boolean
        With frmReport
            .iniciar
            .pdf = App.Path & "\Detalle Envio.pdf"
            .informe = "\EP\rptEPDocumentoPaquete"
            .criterio = "{MENSAJERIAS.CODIGO}=" & DECODIFICADORA.EP_EMPRESAS_MENSAJERIA & " AND {ep_paquetes.ID_PAQUETE}=" & lngPaquete
            .imprimir = False
            .generar
            .Visible = False
        End With
        Unload frmReport
End Function
