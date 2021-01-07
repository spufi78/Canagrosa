VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPEDIDOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPEDIDOS
'***************************************************************
Private ID_PEDIDO As Long
Private ANNO As Long
Private numero As Long
Private Fecha As String
Private total As String
Private ESTADO_ID As Long
Private EMPLEADO_ID As Long
Private OBSERVACIONES As String
Private PROVEEDOR_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPEDIDOS
'***************************************************************
Public Property Let setANNO(ByVal dato As Long)
        ANNO = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        Fecha = dato
End Property
Public Property Let setID_PEDIDO(ByVal dato As Long)
        ID_PEDIDO = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
        numero = dato
End Property
Public Property Let setTOTAL(ByVal dato As String)
        total = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPEDIDOS
'***************************************************************
Public Property Get getANNO() As Long
        getANNO = ANNO
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = Fecha
End Property
Public Property Get getID_PEDIDO() As Long
        getID_PEDIDO = ID_PEDIDO
End Property
Public Property Get getNUMERO() As Long
        getNUMERO = numero
End Property
Public Property Get getTOTAL() As String
        getTOTAL = total
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPEDIDOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PEDIDOS WHERE ID_PEDIDO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ANNO = rs("ANNO")
                ESTADO_ID = rs("ESTADO_ID")
                Fecha = rs("FECHA")
                ID_PEDIDO = rs("ID_PEDIDO")
                numero = rs("NUMERO")
                total = rs("TOTAL")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                OBSERVACIONES = rs("OBSERVACIONES")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPEDIDOS)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PEDIDO) FROM PEDIDOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PEDIDO = 1
        Else
                ID_PEDIDO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Sub Calcular_Numero()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(NUMERO) FROM PEDIDOS WHERE ANNO=" & ANNO
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        numero = 1
    Else
        numero = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Me.Calcular_Numero
        consulta = "INSERT INTO PEDIDOS " & _
                   "  VALUES (" & _
                        ID_PEDIDO & "," & ANNO & "," & numero & ",'" & _
                        Fecha & "','" & total & "'," & ESTADO_ID & "," & _
                        EMPLEADO_ID & ",'" & OBSERVACIONES & "'," & PROVEEDOR_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_PEDIDO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPEDIDOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PEDIDOS SET " & _
                        " ESTADO_ID = " & ESTADO_ID & "," & _
                        " FECHA = '" & Fecha & "'," & _
                        " ANNO = " & ANNO & "," & _
                        " NUMERO = " & numero & "," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " TOTAL = '" & total & "'" & _
                " WHERE ID_PEDIDO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPEDIDOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PEDIDOS_DETALLE " & _
                "    WHERE PEDIDO_ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM PEDIDOS " & _
                "    WHERE ID_PEDIDO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPEDIDOS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_PEDIDO,A.NUMERO, A.ANNO,A.FECHA,B.NOMBRE,C.USUARIO,A.ESTADO_ID,D.NOMBRE,D.ID_PROVEEDOR " & _
                   "  FROM PEDIDOS A, PEDIDOS_ESTADOS B,USUARIOS C,PROVEEDORES D " & _
                   " WHERE A.ESTADO_ID = B.ID_ESTADO " & _
                   "   AND A.PROVEEDOR_ID = D.ID_PROVEEDOR " & _
                   "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
                   " ORDER BY A.ID_PEDIDO DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_pedido(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_PEDIDO,A.NUMERO, A.ANNO,A.FECHA,B.NOMBRE,C.USUARIO,A.ESTADO_ID,D.NOMBRE,D.ID_PROVEEDOR " & _
                   "  FROM PEDIDOS A, PEDIDOS_ESTADOS B,USUARIOS C,PROVEEDORES D " & _
                   " WHERE A.ESTADO_ID = B.ID_ESTADO " & _
                   "   AND A.PROVEEDOR_ID = D.ID_PROVEEDOR " & _
                   "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
                   "   AND A.ID_PEDIDO = " & ID & _
                   " ORDER BY A.ID_PEDIDO DESC"
        Set Listado_por_pedido = datos_bd(consulta)
End Function
Public Function Articulo_en_Pedido(ID As Long) As Integer
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.CANTIDAD " & _
                   "  FROM PEDIDOS A, PEDIDOS_DETALLE B " & _
                   " WHERE B.ARTICULO_ID=" & ID & _
                   "   AND A.ID_PEDIDO = B.PEDIDO_ID " & _
                   "   AND A.ESTADO_ID = 0"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Articulo_en_Pedido = 0
        Else
            Articulo_en_Pedido = rs.Fields(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Articulo_en_Pedido = 0
        MsgBox "Error al cargar los articulos en pedidos (Cargar)", vbCritical, Err.Description
End Function
Public Function Articulo_en_Pedido_Confirmado(ID As Long) As Integer
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.CANTIDAD " & _
                   "  FROM DOCUMENTOS A, DOCUMENTOS_DETALLE B " & _
                   " WHERE B.ARTICULO_ID=" & ID & _
                   "   AND A.ID_DOCUMENTO = B.DOCUMENTO_ID " & _
                   "   AND A.ESTADO_ID = 2"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Articulo_en_Pedido_Confirmado = 0
        Else
            Articulo_en_Pedido_Confirmado = rs.Fields(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Articulo_en_Pedido_Confirmado = 0
        MsgBox "Error al cargar los articulos en pedidos confirmados(Cargar)", vbCritical, Err.Description
End Function


Public Function Listado_Pedido(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT art.id_articulo, " & _
               "       art.ean, " & _
               "       artt.nombre, " & _
               "       art.descripcion, " & _
               "       f.nombre, " & _
               "       s.nombre, " & _
               "       art.precio_venta, " & _
               "       art.stock, " & _
               "       art.tipo_articulo_id, " & _
               "       art.LOTE, " & _
               "       pd.cantidad, " & _
               "       pd.recibido " & _
               " FROM articulos as art, familias as f, subfamilias as s, pedidos_detalle as pd, articulos_tipos as artt " & _
               " WHERE art.familia_id = f.id_familia " & _
               "   AND art.subfamilia_id = s.id_subfamilia " & _
               "   AND art.tipo_articulo_id = artt.id_tipo_articulo " & _
               "   AND pd.pedido_id = " & ID & _
               "   AND pd.articulo_id = art.id_articulo " & _
               " order by pd.orden"
    Set Listado_Pedido = datos_bd(consulta)
End Function
Public Function Articulo_Listado_Pedido(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT art.id_articulo, " & _
               "       art.ean, " & _
               "       artt.nombre, " & _
               "       art.descripcion, " & _
               "       f.nombre, " & _
               "       s.nombre, " & _
               "       art.precio_venta, " & _
               "       art.stock, " & _
               "       art.tipo_articulo_id, " & _
               "       art.LOTE " & _
               " FROM articulos as art, familias as f, subfamilias as s, articulos_tipos as artt " & _
               " WHERE art.familia_id = f.id_familia " & _
               "   AND art.subfamilia_id = s.id_subfamilia " & _
               "   AND art.tipo_articulo_id = artt.id_tipo_articulo " & _
               "   AND art.id_articulo = " & ID
    Set Articulo_Listado_Pedido = datos_bd(consulta)
End Function


Public Function Cambiar_estado(ID As Long, ESTADO As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PEDIDOS " & _
                "      SET ESTADO_ID = " & ESTADO & _
                "    WHERE ID_PEDIDO = " & ID
        execute_bd consulta
        Cambiar_estado = True
        Exit Function
fallo:
        Cambiar_estado = False
        MsgBox "Error al cambiar el estado del pedido (clsPEDIDOS)", vbCritical, Err.Description
End Function


