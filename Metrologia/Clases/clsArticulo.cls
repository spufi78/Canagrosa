VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private ID_ARTICULO As Long
Private FAMILIA_ID As Integer
Private SUBFAMILIA_ID As Integer
Private DESCRIPCION As String
Private EAN As String
Private PRECIO_VENTA As String
Private STOCK As Integer
Private LOTE As Integer
Private PP As Integer
Private SS As Integer
Private VOLUMEN As String
Private TIPO_ARTICULO_ID As Integer
Private ARTICULO_ID_RELACIONADO As Integer
Private IVA As Integer
Public Property Let setID_ARTICULO(ByVal dato As Long)
    ID_ARTICULO = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Integer)
    FAMILIA_ID = dato
End Property
Public Property Let setSUBFAMILIA_ID(ByVal dato As Integer)
    SUBFAMILIA_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Let setEAN(ByVal dato As String)
    EAN = Trim(dato)
End Property
Public Property Let setPRECIO_VENTA(ByVal dato As String)
    PRECIO_VENTA = Trim(dato)
End Property
Public Property Let setSTOCK(ByVal dato As Integer)
    STOCK = dato
End Property
Public Property Let setLOTE(ByVal dato As Integer)
    LOTE = dato
End Property
Public Property Let setPP(ByVal dato As Integer)
    PP = dato
End Property
Public Property Let setSS(ByVal dato As Integer)
    SS = dato
End Property
Public Property Let setVOLUMEN(ByVal dato As String)
    VOLUMEN = dato
End Property
Public Property Let setTIPO_ARTICULO_ID(ByVal dato As Integer)
    TIPO_ARTICULO_ID = dato
End Property
Public Property Let setARTICULO_ID_RELACIONADO(ByVal dato As Integer)
    ARTICULO_ID_RELACIONADO = dato
End Property
Public Property Let setIVA(ByVal dato As Integer)
    IVA = dato
End Property
Public Property Get getID_ARTICULO() As Long
    getID_ARTICULO = ID_ARTICULO
End Property
Public Property Get getFAMILIA_ID() As Integer
    getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getSUBFAMILIA_ID() As Integer
    getSUBFAMILIA_ID = SUBFAMILIA_ID
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getEAN() As String
    getEAN = EAN
End Property
Public Property Get getPRECIO_VENTA() As String
    getPRECIO_VENTA = PRECIO_VENTA
End Property
Public Property Get getSTOCK() As Integer
    getSTOCK = STOCK
End Property
Public Property Get getLOTE() As Integer
    getLOTE = LOTE
End Property
Public Property Get getPP() As Integer
    getPP = PP
End Property
Public Property Get getSS() As Integer
    getSS = SS
End Property
Public Property Get getVOLUMEN() As String
    getVOLUMEN = VOLUMEN
End Property
Public Property Get geTTIPO_ARTICULO_ID() As Integer
    geTTIPO_ARTICULO_ID = TIPO_ARTICULO_ID
End Property
Public Property Get getARTICULO_ID_RELACIONADO() As Integer
    getARTICULO_ID_RELACIONADO = ARTICULO_ID_RELACIONADO
End Property
Public Property Get getIVA() As Integer
    getIVA = IVA
End Property
Public Function Cargar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS WHERE ID_ARTICULO=" & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Cargar = False
        Else
            ID_ARTICULO = rs.Fields("ID_ARTICULO")
            FAMILIA_ID = rs.Fields("FAMILIA_ID")
            SUBFAMILIA_ID = rs.Fields("SUBFAMILIA_ID")
            DESCRIPCION = rs.Fields("DESCRIPCION")
            EAN = rs.Fields("EAN")
            PRECIO_VENTA = rs.Fields("PRECIO_VENTA")
            STOCK = rs.Fields("STOCK")
            LOTE = rs.Fields("LOTE")
            PP = rs.Fields("PP")
            SS = rs.Fields("SS")
            VOLUMEN = rs.Fields("VOLUMEN")
            TIPO_ARTICULO_ID = rs.Fields("TIPO_ARTICULO_ID")
            ARTICULO_ID_RELACIONADO = rs.Fields("ARTICULO_ID_RELACIONADO")
            IVA = rs("IVA")
            Cargar = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Cargar = False
        MsgBox "Error al cargar el artículo (Cargar)", vbCritical, Err.Description
End Function
Public Function Cargar_por_EAN(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ARTICULOS WHERE EAN='" & ID & "'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Cargar_por_EAN = False
        Else
            ID_ARTICULO = rs.Fields("ID_ARTICULO")
            FAMILIA_ID = rs.Fields("FAMILIA_ID")
            SUBFAMILIA_ID = rs.Fields("SUBFAMILIA_ID")
            DESCRIPCION = rs.Fields("DESCRIPCION")
            EAN = rs.Fields("EAN")
            PRECIO_VENTA = rs.Fields("PRECIO_VENTA")
            STOCK = rs.Fields("STOCK")
            LOTE = rs.Fields("LOTE")
            PP = rs.Fields("PP")
            SS = rs.Fields("SS")
            VOLUMEN = rs.Fields("VOLUMEN")
            TIPO_ARTICULO_ID = rs.Fields("TIPO_ARTICULO_ID")
            ARTICULO_ID_RELACIONADO = rs.Fields("ARTICULO_ID_RELACIONADO")
            IVA = rs("IVA")
            Cargar_por_EAN = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Cargar_por_EAN = False
        MsgBox "Error al cargar el artículo (Cargar)", vbCritical, Err.Description
End Function
Public Function Stock_Relacionado(ID As Long) As Integer
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT SUM(STOCK) " & _
                   "  FROM ARTICULOS " & _
                   " WHERE ARTICULO_ID_RELACIONADO = " & ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Stock_Relacionado = 0
        Else
            Stock_Relacionado = rs.Fields(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Stock_Relacionado = 0
        MsgBox "Error al cargar el artículo relacionado(Cargar)", vbCritical, Err.Description
End Function
Public Function Stock_Relacionado_Crudo(ID As Long) As Integer
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT SUM(B.STOCK) " & _
                   "  FROM ARTICULOS A, ARTICULOS B" & _
                   " WHERE A.ARTICULO_ID_RELACIONADO = B.ID_ARTICULO " & _
                   "   AND A.ID_ARTICULO = " & ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Stock_Relacionado_Crudo = 0
        Else
            Stock_Relacionado_Crudo = rs.Fields(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Stock_Relacionado_Crudo = 0
        MsgBox "Error al cargar el artículo relacionado(Cargar)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    If ID_ARTICULO = 0 Then
        Me.CrearId
    End If
    consulta = "INSERT INTO ARTICULOS" & _
               " VALUES(" & _
               ID_ARTICULO & "," & FAMILIA_ID & "," & SUBFAMILIA_ID & ",'" & DESCRIPCION & "','" & _
               EAN & "','" & PRECIO_VENTA & "'," & _
               STOCK & "," & _
               LOTE & "," & PP & "," & SS & ",'" & VOLUMEN & "'," & TIPO_ARTICULO_ID & "," & ARTICULO_ID_RELACIONADO & "," & IVA & ")"
    execute_bd consulta
    ' Insertamos el stock por puesto
'    Dim oArticulo_stock As New clsArticulos_stock
'    With oArticulo_stock
'        .setARTICULO_ID = ID_ARTICULO
'        .setPUESTO_ID = CInt(ReadINI(App.Path + "\config.ini", "Datos", "Puesto"))
'        .setSTOCK = STOCK
'        .Insertar
'    End With
    Insertar = ID_ARTICULO
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Articulo (Insertar)", vbCritical, Err.Description
End Function
Public Function CrearId() As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_ARTICULO) FROM ARTICULOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_ARTICULO = 1
    Else
        ID_ARTICULO = rs.Fields(0) + 1
    End If
    CrearId = ID_ARTICULO
    Set rs = Nothing
End Function
Public Function Modificar(ID As Long) As Boolean
    Dim consulta As String
'    Dim oArticulo_stock As New clsArticulos_stock
    ' Modificamos el stock por puesto
'    With oArticulo_stock
'        If .Carga(ID) = True Then
'            .setSTOCK = STOCK
'            .modificar (ID)
'        Else
'            .setARTICULO_ID = ID
'            .setPUESTO_ID = CInt(ReadINI(App.Path + "\config.ini", "Datos", "Puesto"))
'            .setSTOCK = STOCK
'            .Insertar
'        End If
'    End With
    ' Recuperamos el nuevo total y modificamos el articulo
'    Dim stock_total As Integer
'    stock_total = oArticulo_stock.Total_Articulo(ID)
    On Error GoTo fallo:
    consulta = "UPDATE ARTICULOS SET" & _
               " FAMILIA_ID = " & FAMILIA_ID & "," & _
               " SUBFAMILIA_ID = " & SUBFAMILIA_ID & "," & _
               " DESCRIPCION = '" & DESCRIPCION & "'," & _
               " EAN = '" & EAN & "'," & _
               " PRECIO_VENTA = '" & PRECIO_VENTA & "'," & _
               " STOCK = " & STOCK & "," & _
               " LOTE = " & LOTE & "," & _
               " PP = " & PP & "," & _
               " SS = " & SS & "," & _
               " VOLUMEN = '" & VOLUMEN & "'," & _
               " ARTICULO_ID_RELACIONADO = " & ARTICULO_ID_RELACIONADO & "," & _
               " IVA = " & IVA & "," & _
               " TIPO_ARTICULO_ID = " & TIPO_ARTICULO_ID & _
               " WHERE " & _
               " ID_ARTICULO = " & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el artículo.", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    consulta = "SELECT * FROM DOCUMENTOS_DETALLE WHERE ARTICULO_ID = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "No se puede eliminar. Existen documentos con ese articulo.", vbCritical, App.Title
        Eliminar = False
        Exit Function
    End If
'    consulta = "DELETE FROM ARTICULOS_STOCK " & _
'               " WHERE " & _
'               " ID_ARTICULO = " & ID
'    execute_bd consulta
    consulta = "DELETE FROM ARTICULOS " & _
               " WHERE " & _
               " ID_ARTICULO = " & ID
    execute_bd consulta
    consulta = "DELETE FROM ARTICULOS_PROVEEDORES " & _
               " WHERE " & _
               " ARTICULO_ID = " & ID
    execute_bd consulta
    consulta = "DELETE FROM ARTICULOS_ESCANDALLO " & _
               " WHERE " & _
               " ARTICULO_ID = " & ID
    execute_bd consulta
    consulta = "DELETE FROM ARTICULOS_ESCANDALLO " & _
               " WHERE " & _
               " ARTICULO_ID_ESCANDALLO = " & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el artículo.", vbCritical, Err.Description
End Function
Public Function Listado(des As String, familia As String, subfamilia As String, tipo As String, STOCK As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim sfamilia As String
    Dim ssub As String
    Dim lstock As String
    Dim stipo As String
    If familia <> "" Then
        sfamilia = " AND FAMILIAS.ID_FAMILIA = " & familia
    End If
    If subfamilia <> "" Then
        ssub = " AND SUBFAMILIAS.ID_SUBFAMILIA = " & subfamilia
    End If
    If STOCK = True Then
        lstock = " AND STOCK <= PP "
    End If
    If tipo <> "" Then
        stipo = " AND TIPO_ARTICULO_ID = " & tipo
    End If
    consulta = "SELECT ARTICULOS.id_articulo, " & _
               "       ARTICULOS.ean, " & _
               "       articulos_tipos.nombre, " & _
               "       ARTICULOS.descripcion, " & _
               "       FAMILIAS.nombre, " & _
               "       SUBFAMILIAS.nombre, " & _
               "       ARTICULOS.precio_venta, " & _
               "       ARTICULOS.stock, " & _
               "       ARTICULOS.tipo_articulo_id, " & _
               "       ARTICULOS.FAMILIA_ID, " & _
               "       ARTICULOS.SUBFAMILIA_ID " & _
               " FROM articulos, " & _
               "      familias, " & _
               "      subfamilias, " & _
               "      articulos_tipos " & _
               " WHERE ARTICULOS.familia_id = FAMILIAS.id_familia " & _
               "   AND ARTICULOS.subfamilia_id = SUBFAMILIAS.id_subfamilia " & _
               "   AND ARTICULOS.descripcion like '%" & des & "%'" & _
               "   AND ARTICULOS.tipo_articulo_id = articulos_tipos.id_tipo_articulo " & _
               sfamilia & _
               ssub & _
               lstock & _
               stipo & _
               " order by ARTICULOS.DESCRIPCION"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Faltas(des As String, familia As String, subfamilia As String, tipo As String, criterio As Integer) As ADODB.Recordset
    ' Listado de los productos con stock por debajo del pp y que no son del tipo
    ' Pintados (2)
    Dim consulta As String
    Dim sfamilia As String
    Dim ssub As String
    Dim lstock As String
    Dim stipo As String
    Dim scriterio As String
    If familia <> "" Then
        sfamilia = " AND F.ID_FAMILIA = " & familia
    End If
    If subfamilia <> "" Then
        ssub = " AND S.ID_SUBFAMILIA = " & subfamilia
    End If
    lstock = " AND STOCK <= PP "
    If tipo <> "" Then
        stipo = " AND TIPO_ARTICULO_ID = " & tipo
    End If
    If criterio = -1 Then ' Faltas para pedido a proveedor
        scriterio = " AND TIPO_ARTICULO_ID IN (1,4,5,6) "
    Else ' Faltas para OP
        scriterio = " AND TIPO_ARTICULO_ID = 2 "
    End If
    
    consulta = "SELECT art.id_articulo, " & _
               "       art.ean, " & _
               "       artt.nombre, " & _
               "       art.descripcion, " & _
               "       f.nombre, " & _
               "       s.nombre, " & _
               "       art.precio_venta, " & _
               "       art.stock, " & _
               "       art.tipo_articulo_id, " & _
               "       art.LOTE " & _
               " FROM articulos as art, " & _
               "      familias as f, " & _
               "      subfamilias as s, " & _
               "      articulos_tipos as artt " & _
               " WHERE art.familia_id = f.id_familia " & _
               "   AND art.subfamilia_id = s.id_subfamilia " & _
               "   AND art.descripcion like '%" & des & "%'" & _
               "   AND art.tipo_articulo_id = artt.id_tipo_articulo " & _
               "   AND art.tipo_articulo_id <> 3 " & _
               scriterio & _
               sfamilia & _
               ssub & _
               lstock & _
               stipo & _
               " order by art.DESCRIPCION"
    Set Listado_Faltas = datos_bd(consulta)
End Function
Public Function descontar_unidades(ID As Long, unidades As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
'    consulta = "UPDATE ARTICULOS_STOCK " & _
'               "   SET STOCK = STOCK - " & unidades & _
'               " WHERE " & _
'               "     ARTICULO_ID = " & ID & _
'               "     AND PUESTO_ID = " & CInt(ReadINI(App.Path + "\config.ini", "Datos", "Puesto"))
'    execute_bd consulta
    consulta = "UPDATE ARTICULOS " & _
               "   SET STOCK = STOCK - " & unidades & _
               " WHERE ID_ARTICULO = " & ID
    execute_bd consulta
    descontar_unidades = True
    Exit Function
fallo:
    descontar_unidades = False
'    MsgBox "Error al descontar las unidades del artículo.", vbCritical, Err.Description
End Function
Public Function aumentar_unidades(ID As Long, unidades As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
'    consulta = "UPDATE ARTICULOS_STOCK " & _
'               " SET STOCK = STOCK + " & unidades & _
'               " WHERE " & _
'               " ARTICULO_ID = " & ID & _
'               " AND PUESTO_ID = " & CInt(ReadINI(App.Path + "\config.ini", "Datos", "Puesto"))
'    execute_bd consulta
    consulta = "UPDATE ARTICULOS " & _
               "   SET STOCK = STOCK + " & unidades & _
               " WHERE ID_ARTICULO = " & ID
    execute_bd consulta
    aumentar_unidades = True
    Exit Function
fallo:
    aumentar_unidades = False
'    MsgBox "Error al aumentar las unidades el artículo.", vbCritical, Err.Description
End Function

Public Function duplicado(texto As String) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT id_articulo FROM ARTICULOS where descripcion = '" & texto & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function Cargar_Linea(ARTICULO As Long, TARIFA As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT A.DESCRIPCION,F.NOMBRE,S.NOMBRE,A.STOCK,C.PRECIO_VENTA,A.ID_ARTICULO,A.EAN, A.TIPO_ARTICULO_ID, A.IVA " & _
               " FROM ARTICULOS A, FAMILIAS F, SUBFAMILIAS S, ARTICULOS_TARIFAS C " & _
               " WHERE A.ID_ARTICULO = " & ARTICULO & _
               "   AND A.FAMILIA_ID = F.ID_FAMILIA " & _
               "   AND A.SUBFAMILIA_ID = S.ID_SUBFAMILIA " & _
               "   AND A.ID_ARTICULO = C.ARTICULO_ID " & _
               "   AND C.TARIFA_ID = " & TARIFA
    Set Cargar_Linea = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al buscar el artículo.", vbCritical, Err.Description
End Function
Public Function Cargar_Linea_EAN(codigo As String, TARIFA As Integer) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT A.DESCRIPCION,F.NOMBRE,S.NOMBRE,A.STOCK,C.PRECIO_VENTA,A.ID_ARTICULO,A.EAN, A.TIPO_ARTICULO_ID, A.IVA " & _
               " FROM ARTICULOS A, FAMILIAS F, SUBFAMILIAS S, ARTICULOS_TARIFAS C " & _
               " WHERE A.EAN='" & codigo & "'" & _
               "   AND A.FAMILIA_ID = F.ID_FAMILIA " & _
               "   AND A.SUBFAMILIA_ID = S.ID_SUBFAMILIA " & _
               "   AND A.TIPO_ARTICULO_ID IN (3,4,6,7) " & _
               "   AND A.ID_ARTICULO = C.ARTICULO_ID " & _
               "   AND C.TARIFA_ID = " & TARIFA
    Set Cargar_Linea_EAN = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al buscar el artículo.", vbCritical, Err.Description
End Function
Public Function Codigo_Barras(ARTICULO As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT A.ID_ARTICULO,A.DESCRIPCION,A.PRECIO_VENTA,F.NOMBRE,S.NOMBRE,A.EAN " & _
               " FROM ARTICULOS A, FAMILIAS F, SUBFAMILIAS S " & _
               " WHERE A.ID_ARTICULO=" & ARTICULO & _
               "   AND A.FAMILIA_ID = F.ID_FAMILIA " & _
               "   AND A.SUBFAMILIA_ID = S.ID_SUBFAMILIA"
    Set Codigo_Barras = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al buscar el artículo.", vbCritical, Err.Description
End Function
Public Function Convertir_EAN(codigo As String) As Long
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT ID_ARTICULO " & _
               " FROM ARTICULOS " & _
               " WHERE EAN='" & codigo & "'"
    Convertir_EAN = datos_bd(consulta).Fields(0)
    Exit Function
fallo:
    MsgBox "Error al buscar el artículo.", vbCritical, Err.Description
End Function
Public Function Convertir_a_EAN(ID As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT EAN " & _
               " FROM ARTICULOS " & _
               " WHERE ID_ARTICULO=" & ID
    Convertir_a_EAN = datos_bd(consulta).Fields(0)
    Exit Function
fallo:
    MsgBox "Error al buscar el artículo.", vbCritical, Err.Description
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ARTICULO, DESCRIPCION FROM ARTICULOS " & _
               " WHERE TIPO_ARTICULO_ID = 4 " & _
               " ORDER BY DESCRIPCION"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_por_tipo(tipo As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ARTICULO, DESCRIPCION FROM ARTICULOS " & _
               " WHERE TIPO_ARTICULO_ID = " & tipo & _
               " ORDER BY DESCRIPCION"
    Set Listado_por_tipo = datos_bd(consulta)
End Function

Public Function Listado_Componentes(tipo As String, des As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ARTICULO, DESCRIPCION FROM ARTICULOS " & _
               " WHERE TIPO_ARTICULO_ID IN (" & tipo & ")" & _
               "   AND DESCRIPCION like '%" & des & "%'" & _
               " ORDER BY DESCRIPCION"
    Set Listado_Componentes = datos_bd(consulta)
End Function
Public Function Listado_por_tipo_y_texto(tipo As Integer, des As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ARTICULO, DESCRIPCION,STOCK FROM ARTICULOS " & _
               " WHERE TIPO_ARTICULO_ID = " & tipo & _
               "   AND DESCRIPCION like '%" & des & "%'" & _
               " ORDER BY DESCRIPCION"
    Set Listado_por_tipo_y_texto = datos_bd(consulta)
End Function
Public Function Relacionados(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ARTICULO, DESCRIPCION,STOCK FROM ARTICULOS " & _
               " WHERE ARTICULO_ID_RELACIONADO = " & ID & _
               " ORDER BY DESCRIPCION"
    Set Relacionados = datos_bd(consulta)
End Function
Public Function Listado_para_grid(TARIFA As Integer) As ADODB.Recordset
    ' LP003
    Dim consulta As String
    consulta = "SELECT A.ID_ARTICULO, A.DESCRIPCION, B.PRECIO_VENTA, A.STOCK " & _
               "  FROM ARTICULOS A, ARTICULOS_TARIFAS B " & _
               " WHERE A.ID_ARTICULO = B.ARTICULO_ID " & _
               "   AND B.TARIFA_ID = " & TARIFA & _
               " ORDER BY A.EAN"
    Set Listado_para_grid = datos_bd(consulta)
End Function


