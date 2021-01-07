VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientes_pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCLIENTES_PEDIDOS
'***************************************************************
Private ID_PEDIDO As Long
Private TIPO_ID As Long
Private CODIGO As String
Private DESCRIPCION As String
Private CLIENTE_ID As Long
Private IMPORTE As String
Private FECHA_BAJA As String
Private FECHA_PEDIDO As String
Private FECHA_ALTA As String
Private REVISADO As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCLIENTES_PEDIDOS
'***************************************************************
Public Property Let setID_PEDIDO(ByVal dato As Long)
        ID_PEDIDO = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setFECHA_BAJA(ByVal dato As String)
        FECHA_BAJA = dato
End Property
Public Property Let setFECHA_PEDIDO(ByVal dato As String)
        FECHA_PEDIDO = dato
End Property
Public Property Let setFECHA_ALTA(ByVal dato As String)
        FECHA_ALTA = dato
End Property
Public Property Let setREVISADO(ByVal dato As Integer)
        REVISADO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCLIENTES_PEDIDOS
'***************************************************************
Public Property Get getID_PEDIDO() As Long
        getID_PEDIDO = ID_PEDIDO
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getFECHA_BAJA() As String
        getFECHA_BAJA = FECHA_BAJA
End Property
Public Property Get getFECHA_PEDIDO() As String
        getFECHA_PEDIDO = FECHA_PEDIDO
End Property
Public Property Get getFECHA_ALTA() As String
        getFECHA_ALTA = FECHA_ALTA
End Property
Public Property Get getREVISADO() As Integer
        getREVISADO = REVISADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCLIENTES_PEDIDOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES_PEDIDOS WHERE ID_PEDIDO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PEDIDO = rs("ID_PEDIDO")
                TIPO_ID = rs("TIPO_ID")
                CODIGO = rs("CODIGO")
                DESCRIPCION = rs("DESCRIPCION")
                CLIENTE_ID = rs("CLIENTE_ID")
                IMPORTE = rs("IMPORTE")
                FECHA_BAJA = rs("FECHA_BAJA")
                FECHA_PEDIDO = rs("FECHA_PEDIDO")
                FECHA_ALTA = rs("FECHA_ALTA")
                REVISADO = rs("REVISADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsClientes_pedidos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PEDIDO) FROM CLIENTES_PEDIDOS"
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
        consulta = "INSERT INTO CLIENTES_PEDIDOS " & _
                   "  VALUES (" & _
                        ID_PEDIDO & "," & TIPO_ID & "," & "'" & CODIGO & "'" & "," & "'" & DESCRIPCION & "'" & "," & _
                        CLIENTE_ID & ",'" & IMPORTE & "','" & FECHA_BAJA & "'," & "'" & FECHA_PEDIDO & "',CURRENT_DATE," & REVISADO & _
                ")"
        execute_bd consulta
        Insertar = ID_PEDIDO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsClientes_pedidos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CLIENTES_PEDIDOS SET " & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " IMPORTE = '" & IMPORTE & "'," & _
                        " FECHA_BAJA = '" & FECHA_BAJA & "'," & _
                        " FECHA_PEDIDO = '" & FECHA_PEDIDO & "'," & _
                        " REVISADO = " & REVISADO & _
                " WHERE ID_PEDIDO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsClientes_pedidos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM DOCS_PAGO WHERE PEDIDO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "Existen documentos de pago con ese pedido asociado, no se puede anular.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        consulta = "SELECT * FROM MUESTRAS WHERE PEDIDO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "Existen muestras con ese pedido asociado, no se puede anular.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        consulta = "DELETE FROM CLIENTES_PEDIDOS WHERE ID_PEDIDO = " & ID
        execute_bd consulta
        'M1539-I Eliminar adjuntos del pedido
        Dim oAdjunto As New clsAdjuntos
        oAdjunto.EliminarTodos TOBJETO.TOBJETO_CLIENTES_PEDIDOS, ID
        Set oAdjunto = Nothing
        'M1539-F
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsClientes_pedidos)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long, tipo As String, CODIGO As String, DESCRIPCION As String) As ADODB.Recordset
        Dim consulta As String
        Dim f As String
        If ID <> 0 Then
            f = f & " AND CP.CLIENTE_ID = " & ID
        End If
        If tipo <> "" Then
            If IsNumeric(tipo) Then
                If tipo <> 0 Then
                    f = f & " AND TIPO_ID = " & tipo
                End If
            End If
        End If
        If CODIGO <> "" Then
            f = f & " AND CP.CODIGO LIKE '%" & CODIGO & "%'"
        End If
        If DESCRIPCION <> "" Then
            f = f & " AND CP.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        consulta = "SELECT ID_PEDIDO,CP.CODIGO,CP.DESCRIPCION,FECHA_PEDIDO,FECHA_BAJA,IMPORTE,D.DESCRIPCION,CP.TIPO_ID,CP.FECHA_ALTA " & _
                   " ,IMPORTE - (SELECT IFNULL(sum(d.TOTAL),0) FROM docs_pago d, clientes as cl WHERE d.ANULADO = 0   AND cl.id_cliente = d.cliente_id   AND d.tipo = 2   AND d.pedido_id = CP.ID_PEDIDO) as RESTAN,CP.REVISADO " & _
                   "  FROM CLIENTES_PEDIDOS CP " & _
                   "  LEFT JOIN decodificadora D ON D.CODIGO = " & DECODIFICADORA.PEDIDOS_CLIENTES_TIPOS & " AND CP.TIPO_ID = D.VALOR " & _
                   " WHERE CP.ID_PEDIDO <> 0 " & _
                   f & _
                   " ORDER BY FECHA_PEDIDO DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoCompleto(ID As Long, tipo As String, CODIGO As String, DESCRIPCION As String, fdesde As String, fhasta As String) As ADODB.Recordset
        Dim consulta As String
        Dim f As String
        If ID <> 0 Then
            f = f & " AND CP.CLIENTE_ID = " & ID
        End If
        If tipo <> "" Then
            If IsNumeric(tipo) Then
                If tipo <> 0 Then
                    f = f & " AND TIPO_ID = " & tipo
                End If
            End If
        End If
        If CODIGO <> "" Then
            f = f & " AND CP.CODIGO LIKE '%" & CODIGO & "%'"
        End If
        If DESCRIPCION <> "" Then
            f = f & " AND CP.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        f = f & " AND CP.FECHA_ALTA >= '" & Format(fdesde, "yyyy-mm-dd") & "'"
        f = f & " AND CP.FECHA_ALTA <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
        consulta = "SELECT ID_PEDIDO,CP.CODIGO,CP.DESCRIPCION,FECHA_PEDIDO,FECHA_BAJA,IMPORTE,D.DESCRIPCION,CP.TIPO_ID,CP.FECHA_ALTA, IMPORTE - IFNULL(SUM(DP.TOTAL),0),CLI.NOMBRE " & _
                   "  FROM CLIENTES_PEDIDOS CP " & _
                   "  LEFT JOIN decodificadora D ON D.CODIGO = " & DECODIFICADORA.PEDIDOS_CLIENTES_TIPOS & " AND CP.TIPO_ID = D.VALOR " & _
                   "  LEFT JOIN docs_pago DP ON CP.CLIENTE_ID = DP.CLIENTE_ID_FACTURA AND DP.ANULADO = 0 AND DP.TIPO = 2 AND DP.PEDIDO_ID  = CP.ID_PEDIDO " & _
                   "  LEFT JOIN clientes CLI ON CLI.ID_CLIENTE = CP.CLIENTE_ID " & _
                   " WHERE CP.ID_PEDIDO <> 0 " & _
                   f & _
                   " GROUP BY 1 " & _
                   " ORDER BY FECHA_PEDIDO DESC"
        Set ListadoCompleto = datos_bd(consulta)
End Function

Public Function Listado_Vigor(ID As Long, tipo As String, CODIGO As String, DESCRIPCION As String) As ADODB.Recordset
        Dim consulta As String
        Dim f As String
        If tipo <> "" Then
            If IsNumeric(tipo) Then
                If tipo <> 0 Then
                    f = f & " AND TIPO_ID = " & tipo
                End If
            End If
        End If
        If CODIGO <> "" Then
            f = f & " AND CP.CODIGO LIKE '%" & CODIGO & "%'"
        End If
        If DESCRIPCION <> "" Then
            f = f & " AND CP.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        
        consulta = "SELECT ID_PEDIDO,CP.CODIGO,CP.DESCRIPCION,FECHA_PEDIDO,FECHA_BAJA,IMPORTE,D.DESCRIPCION,CP.TIPO_ID,CP.FECHA_ALTA " & _
                   "  ,IMPORTE - (SELECT IFNULL(sum(d.TOTAL),0) FROM docs_pago d, clientes as cl WHERE d.ANULADO = 0   AND cl.id_cliente = d.cliente_id   AND d.tipo = 2   AND d.pedido_id = CP.ID_PEDIDO) as RESTAN,CP.REVISADO " & _
                   "  FROM CLIENTES_PEDIDOS CP " & _
                   "  LEFT JOIN decodificadora D ON D.CODIGO = " & DECODIFICADORA.PEDIDOS_CLIENTES_TIPOS & " AND CP.TIPO_ID = D.VALOR " & _
                   " WHERE CLIENTE_ID = " & ID & _
                   "   AND FECHA_BAJA >= '" & Format(Date, "yyyy-mm-dd") & "' " & _
                   f & _
                   " ORDER BY FECHA_PEDIDO DESC"
        Set Listado_Vigor = datos_bd(consulta)
End Function
Public Function Listado_completo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PEDIDO,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
                   "  FROM CLIENTES_PEDIDOS " & _
                   " WHERE ID_PEDIDO <> 0 " & _
                   " ORDER BY CODIGO"
        Set Listado_completo = datos_bd(consulta)
End Function
Public Function Listado_por_Cliente(ID As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PEDIDO,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
                   "  FROM CLIENTES_PEDIDOS " & _
                   " WHERE CLIENTE_ID = " & ID & _
                   "   AND ID_PEDIDO <> 0 " & _
                   " ORDER BY CODIGO"
        Set Listado_por_Cliente = datos_bd(consulta)
End Function
Public Function Listado_por_Cliente_en_Vigor(cliente As Long, fdesde As Date, fhasta As Date) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PEDIDO,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
                   "  FROM CLIENTES_PEDIDOS " & _
                   " WHERE CLIENTE_ID = " & cliente & _
                   "   AND ID_PEDIDO <> 0 " & _
                   "   AND FECHA_BAJA >= '" & Format(fhasta, "yyyy-mm-dd") & "' " & _
                   "   AND FECHA_PEDIDO <= '" & Format(fdesde, "yyyy-mm-dd") & "' " & _
                   " ORDER BY CODIGO"
        Set Listado_por_Cliente_en_Vigor = datos_bd(consulta)
End Function
Public Function Listado_en_fecha(ID As Integer, fecha As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PEDIDO,CODIGO,DESCRIPCION,FECHA_PEDIDO,FECHA_BAJA,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
                   "  FROM CLIENTES_PEDIDOS " & _
                   " WHERE CLIENTE_ID = " & ID & _
                   "   AND FECHA_PEDIDO <= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
                   "   AND FECHA_BAJA >= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
                   " ORDER BY FECHA_PEDIDO DESC"
        Set Listado_en_fecha = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PEDIDO,CODIGO FROM CLIENTES_PEDIDOS ORDER BY CODIGO"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Pedido_en_curso(ID As Integer, fecha As String) As Integer
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT MIN(ID_PEDIDO) " & _
                   "  FROM CLIENTES_PEDIDOS " & _
                   " WHERE CLIENTE_ID = " & ID & _
                   "   AND FECHA_PEDIDO <= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
                   "   AND FECHA_BAJA >= '" & Format(fecha, "yyyy-mm-dd") & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            Pedido_en_curso = 0
        Else
            Pedido_en_curso = rs.Fields(0)
        End If
        Set rs = Nothing
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim consulta As String
    consulta = "SELECT ID_PEDIDO,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
               "  FROM CLIENTES_PEDIDOS " & _
               " WHERE ID_PEDIDO <> 0 " & _
               filtro
    
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CLIENTES_PEDIDOS"
        .setDESCRIPCION = "Pedidos"
        .setPK = "ID_PEDIDO"
        .setCAMPO = "CONCAT(CODIGO,' (',DESCRIPCION,')')"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
