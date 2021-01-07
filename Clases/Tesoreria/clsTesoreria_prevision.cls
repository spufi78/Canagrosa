VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTesoreria_prevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTESORERIA_PREVISION
'***************************************************************
Private TOBJETO As Long
Private COBJETO As Long
Private fecha As String
Private IMPORTE As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTESORERIA_PREVISION
'***************************************************************
Public Property Let setTOBJETO(ByVal dato As Long)
        TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As Long)
        COBJETO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTESORERIA_PREVISION
'***************************************************************
Public Property Get getTOBJETO() As Long
        getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As Long
        getCOBJETO = COBJETO
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTESORERIA_PREVISION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TESORERIA_PREVISION WHERE TOBJETO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                fecha = rs("FECHA")
                IMPORTE = rs("IMPORTE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTesoreria_prevision)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TESORERIA_PREVISION " & _
                   "  VALUES (" & _
                        TOBJETO & "," & COBJETO & "," & "'" & fecha & "'" & ",'" & _
                        IMPORTE & "'" & _
                ")"
        execute_bd consulta
        Insertar = TOBJETO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTesoreria_prevision)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TESORERIA_PREVISION SET " & _
                        " COBJETO = " & COBJETO & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " IMPORTE = '" & IMPORTE & "'" & _
                " WHERE TOBJETO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTesoreria_prevision)", vbCritical, Err.Description
End Function
Public Function Eliminar(TOBJETO As Long, COBJETO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TESORERIA_PREVISION " & _
                "    WHERE TOBJETO = " & TOBJETO & _
                "     AND COBJETO = " & COBJETO
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTesoreria_prevision)", vbCritical, Err.Description
End Function
Public Function Listado(TOBJETO As Long, COBJETO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TESORERIA_PREVISION " & _
                   " WHERE TOBJETO = " & TOBJETO & _
                   "   AND COBJETO = " & COBJETO & _
                   " ORDER BY FECHA"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSTESORERIA_PREVISION
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function
Public Function ListadoCobro() As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        consulta = "SELECT 1,'',a.fp_id,b.NOMBRE , sum((a.total - (a.total * a.descuento) /100)+ (((a.total - ((a.total * a.descuento) /100)) * a.IVA) /100)) AS TOTAL " & _
                    "  FROM docs_pago a " & _
                    " inner join clientes c on a.cliente_id = c.id_cliente " & _
                    "  left join formas_pago b on a.FP_ID = b.ID_FP " & _
                    " where a.tipo = " & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA & _
                    "   and a.PAGADO = 0 and a.ANULADO = 0  and year(a.FECHA_FACTURA) >= 2010 " & _
                    "   and c.airbus = 0 " & _
                    "   and a.FECHA_FACTURA  < concat(year(current_date),'-',month(current_date),'-01') " & _
                    " group by a.fp_id " & _
                    " Union " & _
                    " SELECT 2,DATE_FORMAT(a.FECHA_FACTURA,'%m-%Y'),a.fp_id,b.NOMBRE , sum((a.total - (a.total * a.descuento) /100)+ (((a.total - ((a.total * a.descuento) /100)) * a.IVA) /100)) AS TOTAL " & _
                    "  FROM docs_pago a " & _
                    " inner join clientes c on a.cliente_id = c.id_cliente " & _
                    "  left join formas_pago b on a.FP_ID = b.ID_FP " & _
                    " where a.tipo = " & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA & _
                    "   and a.PAGADO = 0 and a.ANULADO = 0  and year(a.FECHA_FACTURA) >= 2010 " & _
                    "   and c.airbus = 0 " & _
                    "   and a.FECHA_FACTURA >= concat(year(current_date),'-',month(current_date),'-01') " & _
                    " group by DATE_FORMAT(a.FECHA_FACTURA,'%m-%Y'),a.fp_id " & _
                    " order by 3,1,2"
        Set ListadoCobro = datos_bd(consulta)
End Function
Public Function ListadoCobroAirbus(pedido As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If pedido = True Then
            filtro = filtro & " and a.pedido_id <> 0 "
        Else
            filtro = filtro & " and a.pedido_id = 0 "
        End If
        consulta = "SELECT 1,'',a.fp_id,b.NOMBRE , sum((a.total - (a.total * a.descuento) /100)+ (((a.total - ((a.total * a.descuento) /100)) * a.IVA) /100)) AS TOTAL " & _
                    "  FROM docs_pago a " & _
                    " inner join clientes c on a.cliente_id = c.id_cliente " & _
                    "  left join formas_pago b on a.FP_ID = b.ID_FP " & _
                    " where a.tipo in (" & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA & "," & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ABONO & ") " & _
                    "   and a.PAGADO = 0 and a.ANULADO = 0  and year(a.FECHA_FACTURA) >= 2010 " & _
                    "   and c.airbus = 1 " & _
                    filtro & _
                    "   and a.FECHA_FACTURA  < concat(year(current_date),'-',month(current_date),'-01') " & _
                    " group by a.fp_id " & _
                    " Union " & _
                    " SELECT 2,DATE_FORMAT(a.FECHA_FACTURA,'%m-%Y'),a.fp_id,b.NOMBRE , sum((a.total - (a.total * a.descuento) /100)+ (((a.total - ((a.total * a.descuento) /100)) * a.IVA) /100)) AS TOTAL " & _
                    "  FROM docs_pago a " & _
                    " inner join clientes c on a.cliente_id = c.id_cliente " & _
                    "  left join formas_pago b on a.FP_ID = b.ID_FP " & _
                    " where a.tipo in (" & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA & "," & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ABONO & ") " & _
                    "   and a.PAGADO = 0 and a.ANULADO = 0  and year(a.FECHA_FACTURA) >= 2010 " & _
                    "   and c.airbus = 1 " & _
                    filtro & _
                    "   and a.FECHA_FACTURA >= concat(year(current_date),'-',month(current_date),'-01') " & _
                    " group by DATE_FORMAT(a.FECHA_FACTURA,'%m-%Y'),a.fp_id " & _
                    " order by 3,1,2"
        Set ListadoCobroAirbus = datos_bd(consulta)
End Function

Public Function ImporteCobroAirbus(fdesde As String, fhasta As String, tipo As Integer) As Single
        Dim consulta As String
        Dim total As Single
        Dim rs As ADODB.Recordset
        Dim filtro As String
        If tipo = 1 Then
            filtro = " and a.pedido_id = 0"
        ElseIf tipo = 2 Then
            filtro = " and a.pedido_id = 1"
        Else
            filtro = " and a.pedido_id = 1"
        End If
        
        
        consulta = "SELECT sum((a.total - (a.total * a.descuento) /100)+ (((a.total - ((a.total * a.descuento) /100)) * a.IVA) /100)) " & _
                " FROM docs_pago a " & _
                " inner join clientes c on a.cliente_id = c.id_cliente " & _
                " left join formas_pago b on a.FP_ID = b.ID_FP " & _
                "where a.tipo = 2 " & _
                "  and a.PAGADO = 0 and a.ANULADO = 0  and year(a.FECHA_FACTURA) >= 2010 " & _
                "  and c.airbus = 1 " & _
                "  and date_add(a.FECHA_FACTURA, INTERVAL b.DIAS DAY) between '" & Format(fdesde, "yyyy-mm-dd") & "' and '" & Format(fhasta, "yyyy-mm-dd") & "'"
        Set rs = datos_bd(consulta)
        total = 0
        If rs.RecordCount > 0 Then
            total = rs(0)
        End If
        ImporteCobroAirbus = total
End Function
Public Function ListadoPdtesPago() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT 1,'',a.FORMAPAGO, b.descripcion, SUM(a.TOTAL) " & _
                " FROM proveedores_facturas a " & _
                " LEFT JOIN decodificadora b ON a.FORMAPAGO = b.VALOR AND b.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                " WHERE a.f_pago = '0000-00-00' AND a.F_VENCIMIENTO < concat(year(current_date),'-',month(current_date),'-01') " & _
                " GROUP BY a.FORMAPAGO " & _
                " Union " & _
                " SELECT 2,DATE_FORMAT(F_VENCIMIENTO,'%m-%Y'),a.FORMAPAGO, b.descripcion, SUM(a.TOTAL) " & _
                " FROM proveedores_facturas a " & _
                " LEFT JOIN decodificadora b ON a.FORMAPAGO = b.VALOR AND b.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                " WHERE a.f_pago = '0000-00-00' AND a.F_VENCIMIENTO >= concat(year(current_date),'-',month(current_date),'-01') " & _
                " GROUP BY DATE_FORMAT(F_VENCIMIENTO,'%m-%Y'),a.FORMAPAGO " & _
                " order by 3,1,2 "
        Set ListadoPdtesPago = datos_bd(consulta)
End Function
Public Function ListadoPagoPrevisto() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT 1,'',a.FORMAPAGO, b.descripcion, SUM(a.TOTAL) " & _
                " FROM proveedores_facturas a " & _
                " LEFT JOIN decodificadora b ON a.FORMAPAGO = b.VALOR AND b.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                " WHERE a.f_pago = '0000-00-00' AND a.F_PREVISTA < concat(year(current_date),'-',month(current_date),'-01') " & _
                " GROUP BY a.FORMAPAGO " & _
                " Union " & _
                " SELECT 2,DATE_FORMAT(F_PREVISTA,'%m-%Y'),a.FORMAPAGO, b.descripcion, SUM(a.TOTAL) " & _
                " FROM proveedores_facturas a " & _
                " LEFT JOIN decodificadora b ON a.FORMAPAGO = b.VALOR AND b.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                " WHERE a.f_pago = '0000-00-00' AND a.F_PREVISTA >= concat(year(current_date),'-',month(current_date),'-01') " & _
                " GROUP BY DATE_FORMAT(F_PREVISTA,'%m-%Y'),a.FORMAPAGO " & _
                " order by 3,1,2 "
        Set ListadoPagoPrevisto = datos_bd(consulta)
End Function

