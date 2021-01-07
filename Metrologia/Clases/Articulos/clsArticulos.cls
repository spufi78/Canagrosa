VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSARTICULOS
'***************************************************************
Private ID_ARTICULO As Long
Private TIPO_ARTICULO_ID As Long
Private DESCRIPCION As String
Private PROVEEDOR_ID As Long
Private PRECIO_COMPRA As String
Private STOCK As String
Private COMENTARIO As String
Private PESO As String
Private COMISION As String
Private INE As Integer
Private INE_CODIGO As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSARTICULOS
'***************************************************************
Public Property Let setID_ARTICULO(ByVal dato As Long)
        ID_ARTICULO = dato
End Property
Public Property Let setTIPO_ARTICULO_ID(ByVal dato As Long)
        TIPO_ARTICULO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setPRECIO_COMPRA(ByVal dato As String)
        PRECIO_COMPRA = dato
End Property
Public Property Let setSTOCK(ByVal dato As String)
        STOCK = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
        COMENTARIO = dato
End Property
Public Property Let setPESO(ByVal dato As String)
        PESO = dato
End Property
Public Property Let setCOMISION(ByVal dato As String)
        COMISION = dato
End Property
Public Property Let setINE(ByVal dato As Integer)
        INE = dato
End Property
Public Property Let setINE_CODIGO(ByVal dato As Long)
        INE_CODIGO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSARTICULOS
'***************************************************************
Public Property Get getID_ARTICULO() As Long
        getID_ARTICULO = ID_ARTICULO
End Property
Public Property Get getTIPO_ARTICULO_ID() As Long
        getTIPO_ARTICULO_ID = TIPO_ARTICULO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getPRECIO_COMPRA() As String
        getPRECIO_COMPRA = PRECIO_COMPRA
End Property
Public Property Get getSTOCK() As String
        getSTOCK = STOCK
End Property
Public Property Get getCOMENTARIO() As String
        getCOMENTARIO = COMENTARIO
End Property
Public Property Get getPESO() As String
        getPESO = PESO
End Property
Public Property Get getCOMISION() As String
        getCOMISION = COMISION
End Property
Public Property Get getINE() As Integer
        getINE = INE
End Property
Public Property Get getINE_CODIGO() As Long
        getINE_CODIGO = INE_CODIGO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSARTICULOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS WHERE ID_ARTICULO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ARTICULO = rs("ID_ARTICULO")
                TIPO_ARTICULO_ID = rs("TIPO_ARTICULO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                PRECIO_COMPRA = rs("PRECIO_COMPRA")
                STOCK = rs("STOCK")
                COMENTARIO = rs("COMENTARIO")
                PESO = rs("PESO")
                COMISION = rs("COMISION")
                INE = rs("INE")
                INE_CODIGO = rs("INE_CODIGO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsArticulos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_ARTICULO) FROM ARTICULOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ARTICULO = 1
        Else
                ID_ARTICULO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_ARTICULO = 0 Then
            Me.CrearID
        End If
        consulta = "INSERT INTO ARTICULOS " & _
                   "  VALUES (" & _
                        ID_ARTICULO & "," & TIPO_ARTICULO_ID & "," & "'" & DESCRIPCION & "'" & "," & _
                        PROVEEDOR_ID & "," & "'" & PRECIO_COMPRA & "'" & "," & "'" & STOCK & "'" & "," & _
                        "'" & COMENTARIO & "'," & _
                        "'" & PESO & "','" & COMISION & "'," & INE & "," & INE_CODIGO & _
                ")"
        execute_bd consulta
        Insertar = ID_ARTICULO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsArticulos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ARTICULOS SET " & _
                        " TIPO_ARTICULO_ID = " & TIPO_ARTICULO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " PRECIO_COMPRA = '" & PRECIO_COMPRA & "'," & _
                        " STOCK = '" & STOCK & "'," & _
                        " PESO = '" & PESO & "'," & _
                        " COMISION = '" & COMISION & "'," & _
                        " INE = " & INE & "," & _
                        " INE_CODIGO = " & INE_CODIGO & "," & _
                        " COMENTARIO = '" & COMENTARIO & "'" & _
                " WHERE ID_ARTICULO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsArticulos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ARTICULOS " & _
                "    WHERE ID_ARTICULO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsArticulos)", vbCritical, Err.Description
End Function
Public Function Listado(DESCRIPCION As String, tipo As String) As ADODB.Recordset
    Dim consulta As String
    Dim stipo As String
    If tipo <> "" Then
        stipo = " AND A.TIPO_ARTICULO_ID = " & tipo
    End If
    consulta = "SELECT A.ID_ARTICULO, " & _
               "       B.NOMBRE, " & _
               "       A.DESCRIPCION, " & _
               "       C.NOMBRE, " & _
               "       A.PRECIO_COMPRA, " & _
               "       A.COMISION " & _
               " FROM ARTICULOS A,ARTICULOS_TIPOS B,PROVEEDORES C " & _
               " WHERE A.TIPO_ARTICULO_ID = B.ID_TIPO_ARTICULO " & _
               "   AND A.PROVEEDOR_ID = C.ID_PROVEEDOR " & _
               "   AND A.DESCRIPCION like '%" & DESCRIPCION & "%'" & _
               stipo & _
               " order by A.DESCRIPCION"
    Set Listado = datos_bd(consulta)
End Function
Public Function ListadoTarifa() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.ID_ARTICULO, " & _
               "       A.DESCRIPCION " & _
               " FROM ARTICULOS A " & _
               " order by A.ID_ARTICULO "
    Set ListadoTarifa = datos_bd(consulta)
End Function

Public Function Listado_PK(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.ID_ARTICULO, " & _
               "       B.NOMBRE, " & _
               "       A.DESCRIPCION, " & _
               "       C.NOMBRE, " & _
               "       A.PRECIO_COMPRA, " & _
               "       A.COMISION " & _
               " FROM ARTICULOS A,ARTICULOS_TIPOS B,PROVEEDORES C " & _
               " WHERE A.TIPO_ARTICULO_ID = B.ID_TIPO_ARTICULO " & _
               "   AND A.PROVEEDOR_ID = C.ID_PROVEEDOR " & _
               "   AND A.ID_ARTICULO = " & ID & _
               " order by A.DESCRIPCION"
    Set Listado_PK = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_ARTICULO,DESCRIPCION FROM ARTICULOS ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
