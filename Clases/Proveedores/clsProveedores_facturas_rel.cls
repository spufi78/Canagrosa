VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_facturas_rel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_FACTURAS_REL
'***************************************************************
Private ID As Long
Private FACTURA_ID As Long
Private TOBJETO As Long
Private COBJETO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_FACTURAS_REL
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setFACTURA_ID(ByVal dato As Long)
        FACTURA_ID = dato
End Property
Public Property Let setTOBJETO(ByVal dato As Long)
        TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As String)
        COBJETO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_FACTURAS_REL
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getFACTURA_ID() As Long
        getFACTURA_ID = FACTURA_ID
End Property
Public Property Get getTOBJETO() As Long
        getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As String
        getCOBJETO = COBJETO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDORES_FACTURAS_REL
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_FACTURAS_REL WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                FACTURA_ID = rs("FACTURA_ID")
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProveedores_facturas_rel)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM PROVEEDORES_FACTURAS_REL"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PROVEEDORES_FACTURAS_REL " & _
                   "  VALUES (" & _
                        ID & "," & FACTURA_ID & "," & TOBJETO & ",'" & _
                        COBJETO & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProveedores_facturas_rel)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_FACTURAS_REL SET " & _
                        " FACTURA_ID = " & FACTURA_ID & "," & _
                        " TOBJETO = " & TOBJETO & "," & _
                        " COBJETO = '" & COBJETO & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProveedores_facturas_rel)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDORES_FACTURAS_REL " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProveedores_facturas_rel)", vbCritical, Err.Description
End Function
Public Function Listado(factura As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT DISTINCT A.ID,A.TOBJETO,A.COBJETO,B.DESCRIPCION, C.CODIGO_SC,C.FECHA_CREACION " & _
                   "  FROM PROVEEDORES_FACTURAS_REL A, decodificadora B, SC_PAQUETES C " & _
                   " WHERE A.FACTURA_ID = " & factura & _
                   "   AND (A.TOBJETO = " & TOBJETO_SC_DETERMINACIONES & _
                   "    OR  A.TOBJETO = " & TOBJETO_SC_EFICACIA & _
                   "    OR  A.TOBJETO = " & TOBJETO_SC_GENERICA & _
                   "    OR  A.TOBJETO = " & TOBJETO_SC_PEACH & ")" & _
                   "   AND A.COBJETO = C.ID_PAQUETE " & _
                   "   AND A.TOBJETO = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_TOBJETOS & _
                   "   AND C.ESTADO <> " & SC_ESTADO_HISTORICO
        consulta = consulta & " UNION "
        consulta = consulta & "SELECT DISTINCT A.ID,A.TOBJETO,A.COBJETO,B.DESCRIPCION, A.COBJETO,C.FECHA_PEDIDO " & _
                   "  FROM PROVEEDORES_FACTURAS_REL A, decodificadora B, PEDIDOS_BOTE_EX C " & _
                   " WHERE A.FACTURA_ID = " & factura & _
                   "   AND A.TOBJETO = " & TOBJETO_PEDIDO_BOTE_EX & _
                   "   AND A.TOBJETO = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_TOBJETOS & _
                   "   AND LEFT(A.COBJETO,4) = C.CODIGO_PEDIDO_PROVEEDOR " & _
                   "   AND RIGHT(A.COBJETO,4) = YEAR(C.FECHA_PEDIDO) "
        consulta = consulta & " UNION "
        consulta = consulta & "SELECT DISTINCT A.ID,A.TOBJETO,A.COBJETO,B.DESCRIPCION, CONCAT(C.NUMERO,'/',C.ANNO),C.FECHA_CREACION " & _
                   "  FROM PROVEEDORES_FACTURAS_REL A, decodificadora B, PP C " & _
                   " WHERE A.FACTURA_ID = " & factura & _
                   "   AND A.TOBJETO = " & TOBJETO_PEDIDO_PROVEEDOR & _
                   "   AND A.TOBJETO = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_TOBJETOS & _
                   "   AND A.COBJETO = C.ID_PP " & _
                   " ORDER BY ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPROVEEDORES_FACTURAS_REL
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,FACTURA_ID FROM PROVEEDORES_FACTURAS_REL ORDER BY FACTURA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
