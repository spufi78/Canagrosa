VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulos_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private ARTICULO_ID As Long
Private PROVEEDOR_ID As Integer
Private CODIGO_PROVEEDOR As String
Private PRECIO_COSTO As String
Private IVA As Integer
Public Property Let setARTICULO_ID(ByVal dato As Long)
    ARTICULO_ID = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Integer)
    PROVEEDOR_ID = dato
End Property
Public Property Let setCODIGO_PROVEEDOR(ByVal dato As String)
    CODIGO_PROVEEDOR = Trim(dato)
End Property
Public Property Let setPRECIO_COSTO(ByVal dato As String)
    PRECIO_COSTO = Trim(dato)
End Property
Public Property Let setIVA(ByVal dato As Integer)
    IVA = dato
End Property
Public Property Get getARTICULO_ID() As Long
    getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getPROVEEDOR_ID() As Integer
    getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getCODIGO_PROVEEDOR() As String
    getCODIGO_PROVEEDOR = CODIGO_PROVEEDOR
End Property
Public Property Get getPRECIO_COSTO() As String
    getPRECIO_COSTO = PRECIO_COSTO
End Property
Public Property Get getIVA() As Integer
    getIVA = IVA
End Property
Public Function Cargar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS_PROVEEDORES WHERE ARTICULO_ID=" & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Cargar = False
        Else
            ARTICULO_ID = rs.Fields("ARTICULO_ID")
            PROVEEDOR_ID = rs.Fields("PROVEEDOR_ID")
            CODIGO_PROVEEDOR = rs.Fields("CODIGO_PROVEEDOR")
            PRECIO_COSTO = rs.Fields("PRECIO_COSTO")
            IVA = rs.Fields("IVA")
            Cargar = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Cargar = False
        MsgBox "Error al cargar el artículo (Cargar)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "INSERT INTO ARTICULOS_PROVEEDORES" & _
               " VALUES(" & _
               ARTICULO_ID & "," & PROVEEDOR_ID & ",'" & CODIGO_PROVEEDOR & "','" & _
               PRECIO_COSTO & "'," & IVA & ")"
    execute_bd consulta
    Insertar = ARTICULO_ID
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Articulo (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "DELETE FROM ARTICULOS_PROVEEDORES " & _
               " WHERE " & _
               " ARTICULO_ID = " & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el artículo.", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT P.NOMBRE,A.CODIGO,A.PRECIO_COSTO,A.IVA,A.PROVEEDOR_ID " & _
                   "  FROM ARTICULOS_PROVEEDORES A,PROVEEDORES P " & _
                   " WHERE ARTICULO_ID = " & ID & _
                   "   AND A.PROVEEDOR_ID = P.ID_PROVEEDOR"
        Set Listado = datos_bd(consulta)
End Function

Public Function Proveedor_mas_barato(ID As Long) As Integer
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT PROVEEDOR_ID FROM ARTICULOS_PROVEEDORES WHERE ARTICULO_ID=" & ID & " ORDER BY PRECIO_COSTO ASC"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Proveedor_mas_barato = 0
        Else
            Proveedor_mas_barato = rs(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Proveedor_mas_barato = 0
End Function
