VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPEDIDOS_DETALLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPEDIDOS_DETALLE
'***************************************************************
Private PEDIDO_ID As Long
Private ORDEN As Long
Private ARTICULO_ID As Long
Private EAN As String
Private CANTIDAD As Long
Private RECIBIDO As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPEDIDOS_DETALLE
'***************************************************************
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Long)
        CANTIDAD = dato
End Property
Public Property Let setEAN(ByVal dato As String)
        EAN = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setRECIBIDO(ByVal dato As Long)
        RECIBIDO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPEDIDOS_DETALLE
'***************************************************************
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getCANTIDAD() As Long
        getCANTIDAD = CANTIDAD
End Property
Public Property Get getEAN() As String
        getEAN = EAN
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getRECIBIDO() As Long
        getRECIBIDO = RECIBIDO
End Property
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PEDIDOS_DETALLE " & _
                   "  VALUES (" & _
                        PEDIDO_ID & "," & ORDEN & "," & ARTICULO_ID & ",'" & EAN & "'," & _
                        CANTIDAD & "," & RECIBIDO & _
                ")"
        execute_bd consulta
        Insertar = PEDIDO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPEDIDOS_DETALLE)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PEDIDOS_DETALLE " & _
                "    WHERE PEDIDO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPEDIDOS_DETALLE)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PEDIDOS_DETALLE WHERE PEDIDO_ID = " & ID & " ORDER BY ARTICULO_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function recibir_articulo(PEDIDO As Long, articulo As Long, CANTIDAD As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PEDIDOS_DETALLE " & _
                "      SET RECIBIDO = RECIBIDO + " & CANTIDAD & _
                "    WHERE PEDIDO_ID = " & PEDIDO & _
                "      AND ARTICULO_ID = " & articulo
        execute_bd consulta
        recibir_articulo = True
        Exit Function
fallo:
        recibir_articulo = False
        MsgBox "Error al recibir el artículo (clsPEDIDOS_DETALLE)", vbCritical, Err.Description
End Function

