VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuministros_lotes_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSUMINISTROS_LOTES_CLIENTES
'***************************************************************
Private LOTE_ID As Long
Private ORDEN As Long
Private CLIENTE_ID As Long
Private CAPACIDAD_ID As Long
Private NUMERO_BOTES As Long
Private PRECIO As String
Private PEDIDO_ID As Long
Private DOC_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSUMINISTROS_LOTES_CLIENTES
'***************************************************************
Public Property Let setLOTE_ID(ByVal dato As Long)
        LOTE_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setCAPACIDAD_ID(ByVal dato As Long)
        CAPACIDAD_ID = dato
End Property
Public Property Let setNUMERO_BOTES(ByVal dato As Long)
        NUMERO_BOTES = dato
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
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSUMINISTROS_LOTES_CLIENTES
'***************************************************************
Public Property Get getLOTE_ID() As Long
        getLOTE_ID = LOTE_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getCAPACIDAD_ID() As Long
        getCAPACIDAD_ID = CAPACIDAD_ID
End Property
Public Property Get getNUMERO_BOTES() As Long
        getNUMERO_BOTES = NUMERO_BOTES
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
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSUMINISTROS_LOTES_CLIENTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SUMINISTROS_LOTES_CLIENTES WHERE LOTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                LOTE_ID = rs("LOTE_ID")
                ORDEN = rs("ORDEN")
                CLIENTE_ID = rs("CLIENTE_ID")
                CAPACIDAD_ID = rs("CAPACIDAD_ID")
                NUMERO_BOTES = rs("NUMERO_BOTES")
                PRECIO = rs("PRECIO")
                PEDIDO_ID = rs("PEDIDO_ID")
                DOC_ID = rs("DOC_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSuministros_lotes_clientes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO SUMINISTROS_LOTES_CLIENTES " & _
                   "  VALUES (" & _
                        LOTE_ID & "," & ORDEN & "," & CLIENTE_ID & "," & _
                        CAPACIDAD_ID & "," & NUMERO_BOTES & ",'" & PRECIO & "'," & PEDIDO_ID & "," & DOC_ID & _
                ")"
        execute_bd consulta
        Insertar = LOTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSuministros_lotes_clientes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SUMINISTROS_LOTES_CLIENTES SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " CAPACIDAD_ID = " & CAPACIDAD_ID & "," & _
                        " NUMERO_BOTES = " & NUMERO_BOTES & "," & _
                        " PRECIO = '" & PRECIO & "'," & _
                        " PEDIDO_ID = " & PEDIDO_ID & "," & _
                        " DOC_ID = " & DOC_ID & "" & _
                " WHERE LOTE_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSuministros_lotes_clientes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SUMINISTROS_LOTES_CLIENTES " & _
                "    WHERE LOTE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSuministros_lotes_clientes)", vbCritical, Err.Description
End Function
Public Function Listado(lote As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.NOMBRE,C.DESCRIPCION,A.NUMERO_BOTES,A.CLIENTE_ID,A.CAPACIDAD_ID,A.PRECIO,CONCAT(D.CODIGO,' ',D.DESCRIPCION) AS PEDIDO,A.PEDIDO_ID,A.DOC_ID " & _
                   "  FROM SUMINISTROS_LOTES_CLIENTES A " & _
                   " INNER JOIN CLIENTES B ON A.CLIENTE_ID = B.ID_CLIENTE " & _
                   " INNER JOIN ALODINE_CAPACIDAD C ON A.CAPACIDAD_ID = C.ID_CAPACIDAD " & _
                   " LEFT JOIN CLIENTES_PEDIDOS D ON A.PEDIDO_ID = D.ID_PEDIDO " & _
                   " WHERE A.LOTE_ID = " & lote & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function doc_pago(lotes As String, cliente As Long, DOC As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SUMINISTROS_LOTES_CLIENTES " & _
                   "   SET DOC_ID = " & DOC & _
                   " WHERE LOTE_ID IN (" & lotes & ")" & _
                   "   AND CLIENTE_ID = " & cliente & _
                   "   AND DOC_ID = 0"
        execute_bd consulta
        doc_pago = True
        Exit Function
fallo:
        doc_pago = False
        MsgBox "Error al informar el documento de pago (clsAlodine_planificacion)", vbCritical, Err.Description
End Function
Public Function doc_pago_pedido(lote As String, cliente As Long, pedido As Long, DOC As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SUMINISTROS_LOTES_CLIENTES " & _
                   "   SET DOC_ID = " & DOC & _
                   " WHERE CLIENTE_ID = " & cliente & _
                   "   AND PEDIDO_ID = " & pedido & _
                   "   and LOTE_ID in (" & lote & ") " & _
                   "   AND DOC_ID = 0"
        execute_bd consulta
        doc_pago_pedido = True
        Exit Function
fallo:
        doc_pago_pedido = False
        MsgBox "Error al informar el documento de pago (doc_pago_pedido)", vbCritical, Err.Description
End Function


