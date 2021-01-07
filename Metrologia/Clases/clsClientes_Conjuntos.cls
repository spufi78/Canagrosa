VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientes_Conjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCLIENTES_CONJUNTOS
'***************************************************************
Private CLIENTE_ID As Long
Private ARTICULO_ID As Long
Private PRECIO_VENTA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCLIENTES_CONJUNTOS
'***************************************************************
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setPRECIO_VENTA(ByVal dato As String)
        PRECIO_VENTA = dato
End Property
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCLIENTES_CONJUNTOS
'***************************************************************
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getPRECIO_VENTA() As String
        getPRECIO_VENTA = PRECIO_VENTA
End Property
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCLIENTES_CONJUNTOS
'***************************************************************
Public Function Carga(CLIENTE As Long, ARTICULO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES_CONJUNTOS " & _
                   " WHERE CLIENTE_ID = " & CLIENTE & _
                   "   AND ARTICULO_ID = " & ARTICULO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PRECIO_VENTA = rs("PRECIO_VENTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(CLSCLIENTES_CONJUNTOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CLIENTES_CONJUNTOS " & _
                   "  VALUES (" & _
                        CLIENTE_ID & "," & ARTICULO_ID & ",'" & PRECIO_VENTA & "'" & _
                ")"
        conn.Execute consulta
        Insertar = CLIENTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (CLSCLIENTES_CONJUNTOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ARTICULO As Long, TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CLIENTES_CONJUNTOS SET " & _
                        " PRECIO_VENTA = '" & PRECIO_VENTA & "'" & _
                   " WHERE CLIENTE_ID = " & ARTICULO & _
                   "   AND ARTICULO_ID = " & TARIFA
        conn.Execute consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (CLSCLIENTES_CONJUNTOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CLIENTES_CONJUNTOS " & _
                "    WHERE CLIENTE_ID = " & ID
        conn.Execute consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (CLSCLIENTES_CONJUNTOS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES_CONJUNTOS ORDER BY CLIENTE_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_cliente(CLIENTE As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES_CONJUNTOS " & _
                   " WHERE CLIENTE_ID = " & CLIENTE & _
                   " ORDER BY ARTICULO_ID"
        Set Listado_por_cliente = datos_bd(consulta)
End Function


