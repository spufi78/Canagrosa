VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulos_stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsArticulos_stock
'***************************************************************
Private ARTICULO_ID As Long
Private PUESTO_ID As Long
Private STOCK As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsArticulos_stock
'***************************************************************
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setPUESTO_ID(ByVal dato As Long)
        PUESTO_ID = dato
End Property
Public Property Let setSTOCK(ByVal dato As Long)
        STOCK = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsArticulos_stock
'***************************************************************
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getPUESTO_ID() As Long
        getPUESTO_ID = PUESTO_ID
End Property
Public Property Get getSTOCK() As Long
        getSTOCK = STOCK
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsArticulos_stock
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_STOCK " & _
                   " WHERE ARTICULO_ID = " & ID & _
                   "   AND PUESTO_ID = " & CInt(ReadINI(App.Path + "\config.ini", "Datos", "Puesto"))
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ARTICULO_ID = rs("ARTICULO_ID")
                PUESTO_ID = rs("PUESTO_ID")
                STOCK = rs("STOCK")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsArticulos_stock)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ARTICULOS_STOCK " & _
                   "  VALUES (" & _
                        ARTICULO_ID & "," & PUESTO_ID & "," & STOCK & _
                ")"
        conn.Execute consulta
        Insertar = ARTICULO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsArticulos_stock)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ARTICULOS_STOCK SET " & _
                   " STOCK = " & STOCK & "" & _
                   " WHERE ARTICULO_ID = " & ID & _
                   "   AND PUESTO_ID = " & CInt(ReadINI(App.Path + "\config.ini", "Datos", "Puesto"))
        conn.Execute consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsArticulos_stock)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ARTICULOS_STOCK " & _
                "    WHERE ARTICULO_ID = " & ID
        conn.Execute consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsArticulos_stock)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_STOCK ORDER BY ARTICULO_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Total_Articulo(ID As Integer) As Integer
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT SUM(STOCK) FROM ARTICULOS_STOCK " & _
                   " WHERE ARTICULO_ID = " & ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Total_Articulo = 0
        Else
            Total_Articulo = rs.Fields(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Total_Articulo = False
        MsgBox "Error al cargar el total del artículo(clsArticulos_stock)", vbCritical, Err.Description
End Function
Public Function Stock_Tienda(ID As Integer) As Integer
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT SUM(STOCK) FROM ARTICULOS_STOCK " & _
                   " WHERE ARTICULO_ID = " & ID & _
                   "   AND PUESTO_ID = " & CInt(ReadINI(App.Path + "\config.ini", "Datos", "Puesto"))
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Stock_Tienda = 0
        Else
            Stock_Tienda = rs.Fields(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Stock_Tienda = False
        MsgBox "Error al cargar el total del artículo(clsArticulos_stock)", vbCritical, Err.Description
End Function
Public Function Listado_por_tienda(ID As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.DESCRIPCION, P.DESCRIPCION, AST.STOCK, A.STOCK " & _
                   "  FROM ARTICULOS A,ARTICULOS_STOCK AST, PUESTOS P " & _
                   " WHERE AST.PUESTO_ID = P.ID_PUESTO " & _
                   "   AND AST.ARTICULO_ID = A.ID_ARTICULO " & _
                   "   AND A.ID_ARTICULO = " & ID
        Set Listado_por_tienda = datos_bd(consulta)
End Function

