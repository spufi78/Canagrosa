VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSALODINE_CLIENTES
'***************************************************************
Private ALODINE_ID As Long
Private CLIENTE_ID As Long
Private EADS As Long
Private CAPACIDAD_ID As Long
Private ETIQUETA_ID As Long
Private NUMERO_BOTES As Long
Private PRECIO As String
Private pedido As String
Private PEDIDO_ID As Long
Private NORMA As String
Private NORMA_ETIQUETA As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSALODINE_CLIENTES
'***************************************************************
Public Property Let setALODINE_ID(ByVal dato As Long)
        ALODINE_ID = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setEADS(ByVal dato As Long)
        EADS = dato
End Property
Public Property Let setCAPACIDAD_ID(ByVal dato As Long)
        CAPACIDAD_ID = dato
End Property
Public Property Let setETIQUETA_ID(ByVal dato As Long)
        ETIQUETA_ID = dato
End Property
Public Property Let setNUMERO_BOTES(ByVal dato As Long)
        NUMERO_BOTES = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setPEDIDO(ByVal dato As String)
        pedido = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setNORMA(ByVal dato As String)
        NORMA = dato
End Property
Public Property Let setNORMA_ETIQUETA(ByVal dato As Integer)
        NORMA_ETIQUETA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSALODINE_CLIENTES
'***************************************************************
Public Property Get getALODINE_ID() As Long
        getALODINE_ID = ALODINE_ID
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getEADS() As Long
        getEADS = EADS
End Property
Public Property Get getCAPACIDAD_ID() As Long
        getCAPACIDAD_ID = CAPACIDAD_ID
End Property
Public Property Get getETIQUETA_ID() As Long
        getETIQUETA_ID = ETIQUETA_ID
End Property
Public Property Get getNUMERO_BOTES() As Long
        getNUMERO_BOTES = NUMERO_BOTES
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getPEDIDO() As String
        getPEDIDO = pedido
End Property
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getNORMA() As String
        getNORMA = NORMA
End Property
Public Property Get getNORMA_ETIQUETA() As Integer
        getNORMA_ETIQUETA = NORMA_ETIQUETA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSALODINE_CLIENTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_CLIENTES WHERE ALODINE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ALODINE_ID = rs("ALODINE_ID")
                CLIENTE_ID = rs("CLIENTE_ID")
                EADS = rs("EADS")
                CAPACIDAD_ID = rs("CAPACIDAD_ID")
                ETIQUETA_ID = rs("ETIQUETA_ID")
                NUMERO_BOTES = rs("NUMERO_BOTES")
                PRECIO = rs("PRECIO")
                pedido = rs("PEDIDO")
                PEDIDO_ID = rs("PEDIDO_ID")
                NORMA = rs("NORMA")
                NORMA_ETIQUETA = rs("NORMA_ETIQUETA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine_clientes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ALODINE_CLIENTES " & _
                   "  VALUES (" & _
                        ALODINE_ID & "," & CLIENTE_ID & "," & EADS & "," & _
                        CAPACIDAD_ID & "," & ETIQUETA_ID & "," & NUMERO_BOTES & ",'" & _
                        PRECIO & "','" & pedido & "'," & PEDIDO_ID & ",'" & NORMA & "'," & NORMA_ETIQUETA & _
                ")"
        execute_bd consulta
        Insertar = ALODINE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAlodine_clientes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_CLIENTES SET " & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " EADS = " & EADS & "," & _
                        " CAPACIDAD_ID = " & CAPACIDAD_ID & "," & _
                        " ETIQUETA_ID = " & ETIQUETA_ID & "," & _
                        " NUMERO_BOTES = " & NUMERO_BOTES & "," & _
                        " PRECIO = '" & PRECIO & "'," & _
                        " NORMA = '" & NORMA & "'," & _
                        " NORMA_ETIQUETA = " & NORMA_ETIQUETA & "," & _
                        " PEDIDO = '" & pedido & "'," & _
                        " PEDIDO_ID = " & PEDIDO_ID & _
                " WHERE ALODINE_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAlodine_clientes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_CLIENTES " & _
                "    WHERE ALODINE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAlodine_clientes)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_CLIENTES WHERE ALODINE_ID = " & ID & " ORDER BY ALODINE_ID, CLIENTE_ID "
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Clientes(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT C.NOMBRE,AC.DESCRIPCION,AE.DESCRIPCION,A.NUMERO_BOTES,A.PRECIO,A.CLIENTE_ID," & _
                   "       A.CAPACIDAD_ID,A.ETIQUETA_ID,A.EADS,A.CLIENTE_ID,A.PEDIDO,0,A.NORMA,'',A.NORMA_ETIQUETA,'',A.PEDIDO_ID " & _
                   "  FROM ALODINE_CLIENTES A " & _
                   " INNER JOIN CLIENTES C ON A.CLIENTE_ID = C.ID_CLIENTE " & _
                   " INNER JOIN ALODINE_CAPACIDAD AC ON A.CAPACIDAD_ID = AC.ID_CAPACIDAD " & _
                   " INNER JOIN ALODINE_ETIQUETAS AE ON A.ETIQUETA_ID = AE.ID_ETIQUETA " & _
                   " WHERE A.ALODINE_ID = " & ID & _
                   " ORDER BY C.NOMBRE"
        Set Listado_Clientes = datos_bd(consulta)
End Function

