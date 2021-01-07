VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSALODINE_LOTES
'***************************************************************
Private ID_LOTE As Long
Private ALODINE_ID As Long
Private NUMERO_LOTE As Long
Private MUESTRA_ID_EC As Long
Private muestras As String
Private FECHA_ALTA As String
Private fecha_caducidad As String
Private fecha_impresion As String
Private EDICION As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSALODINE_LOTES
'***************************************************************
Public Property Let setID_LOTE(ByVal dato As Long)
        ID_LOTE = dato
End Property
Public Property Let setALODINE_ID(ByVal dato As Long)
        ALODINE_ID = dato
End Property
Public Property Let setNUMERO_LOTE(ByVal dato As Long)
        NUMERO_LOTE = dato
End Property
Public Property Let setMUESTRA_ID_EC(ByVal dato As Long)
        MUESTRA_ID_EC = dato
End Property
Public Property Let setMUESTRAS(ByVal dato As String)
        muestras = dato
End Property
Public Property Let setFECHA_ALTA(ByVal dato As String)
        FECHA_ALTA = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
        fecha_caducidad = dato
End Property
Public Property Let setfecha_impresion(ByVal dato As String)
        fecha_impresion = dato
End Property
Public Property Let setEDICION(ByVal dato As Integer)
        EDICION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSALODINE_LOTES
'***************************************************************
Public Property Get getID_LOTE() As Long
        getID_LOTE = ID_LOTE
End Property
Public Property Get getALODINE_ID() As Long
        getALODINE_ID = ALODINE_ID
End Property
Public Property Get getNUMERO_LOTE() As Long
        getNUMERO_LOTE = NUMERO_LOTE
End Property
Public Property Get getMUESTRA_ID_EC() As Long
        getMUESTRA_ID_EC = MUESTRA_ID_EC
End Property
Public Property Get getMUESTRAS() As String
        getMUESTRAS = muestras
End Property
Public Property Get getFECHA_ALTA() As String
        getFECHA_ALTA = FECHA_ALTA
End Property
Public Property Get getFECHA_CADUCIDAD() As String
        getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getfecha_impresion() As String
        getfecha_impresion = fecha_impresion
End Property
Public Property Get getEDICION() As Integer
        getEDICION = EDICION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSALODINE_LOTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_LOTES WHERE ID_LOTE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_LOTE = rs("ID_LOTE")
                ALODINE_ID = rs("ALODINE_ID")
                NUMERO_LOTE = rs("NUMERO_LOTE")
                MUESTRA_ID_EC = rs("MUESTRA_ID_EC")
                muestras = rs("MUESTRAS")
                FECHA_ALTA = rs("FECHA_ALTA")
                fecha_caducidad = rs("FECHA_CADUCIDAD")
                fecha_impresion = rs("fecha_impresion")
                EDICION = rs("EDICION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine_lotes)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_LOTE) FROM ALODINE_LOTES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_LOTE = 1
        Else
                ID_LOTE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Dim idMuestraEc As Long
        idMuestraEc = Me.recuperarMuestraEC(ALODINE_ID)
        
        consulta = "INSERT INTO ALODINE_LOTES " & _
                   "  VALUES (" & _
                        ID_LOTE & "," & ALODINE_ID & "," & NUMERO_LOTE & "," & idMuestraEc & ",'" & muestras & "','" & _
                        FECHA_ALTA & "'" & "," & "'" & fecha_caducidad & "','" & fecha_impresion & "',1" & _
                ")"
        execute_bd consulta
        Insertar = ID_LOTE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAlodine_lotes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim idMuestraEc As Long
        idMuestraEc = Me.recuperarMuestraEC(ALODINE_ID)
        
        consulta = "UPDATE ALODINE_LOTES SET " & _
                        " NUMERO_LOTE = '" & NUMERO_LOTE & "'," & _
                        " MUESTRA_ID_EC = " & idMuestraEc & "," & _
                        " fecha_impresion = '" & fecha_impresion & "'," & _
                        " MUESTRAS = '" & muestras & "'," & _
                        " FECHA_ALTA = '" & FECHA_ALTA & "'," & _
                        " FECHA_CADUCIDAD = '" & fecha_caducidad & "'" & _
                " WHERE ID_LOTE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAlodine_lotes)", vbCritical, Err.Description
End Function
Public Function aumentarEdicion(ID_LOTE As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_LOTES SET EDICION = EDICION + 1 " & _
                   " WHERE ID_LOTE = " & ID_LOTE
        execute_bd consulta
        aumentarEdicion = True
        Exit Function
fallo:
        aumentarEdicion = False
        MsgBox "Error al aumentarEdicion (clsAlodine_lotes)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_PLANIFICACION " & _
                "    WHERE LOTE_ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM ALODINE_RESULTADOS " & _
                "    WHERE LOTE_ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM ALODINE_LOTES " & _
                "    WHERE ID_LOTE = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAlodine_lotes)", vbCritical, Err.Description
End Function
Public Function Listado(fdesde As String, fhasta As String, NUMERO As String, CODIGO As String, producto As String, CLIENTE_ID As Long) As ADODB.Recordset
        Dim consulta As String
        Dim q As String
        If NUMERO <> "" Then
            If IsNumeric(NUMERO) Then
                q = " AND AL.NUMERO_LOTE = " & NUMERO
            End If
        End If
        If producto <> "" Then
            q = q & " AND A.PRODUCTO LIKE '%" & producto & "%'"
        End If
        If CLIENTE_ID <> 0 Then
            q = q & " AND AP.CLIENTE_ID = " & CLIENTE_ID
        End If
        consulta = "SELECT DISTINCT A.PRODUCTO,A.CODIGO, A.LOTE,A.MADRE,AL.NUMERO_LOTE,AL.FECHA_ALTA,AL.FECHA_CADUCIDAD,AL.ID_LOTE,A.DESCRIPCION " & _
                   "  FROM ALODINE_LOTES AL " & _
                   " INNER JOIN ALODINE A ON AL.ALODINE_ID = A.ID_ALODINE " & _
                   "  LEFT JOIN ALODINE_PLANIFICACION AP ON AL.ID_LOTE = AP.LOTE_ID " & _
                   " WHERE AL.FECHA_ALTA BETWEEN '" & Format(fdesde, "yyyy-mm-dd") & "' AND '" & Format(fhasta, "yyyy-mm-dd") & "' " & _
                   "   AND A.CODIGO LIKE '%" & CODIGO & "%'" & _
                   q & _
                   " ORDER BY ID_LOTE DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Pendientes_Facturar(CLIENTE_ID As Long, fdesde As Date, fhasta As Date) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If CLIENTE_ID <> 0 Then
            filtro = filtro & " AND AP.CLIENTE_ID = " & CLIENTE_ID
        End If
        filtro = filtro & " AND AL.FECHA_ALTA >= '" & Format(fdesde, "yyyy-mm-dd") & "'"
        filtro = filtro & " AND AL.FECHA_ALTA <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
    
        consulta = "SELECT C.NOMBRE,A.PRODUCTO,A.CODIGO,AL.NUMERO_LOTE,AL.FECHA_ALTA,AL.FECHA_CADUCIDAD,AL.ID_LOTE,A.DESCRIPCION,C.ID_CLIENTE,AP.PEDIDO_ID,CP.CODIGO,AP.EADS " & _
                   "  FROM ALODINE_LOTES AL " & _
                   " INNER JOIN ALODINE A on AL.ALODINE_ID = A.ID_ALODINE " & _
                   " INNER JOIN ALODINE_PLANIFICACION AP on AL.ID_LOTE = AP.LOTE_ID " & _
                   " INNER JOIN CLIENTES C ON AP.CLIENTE_ID = C.ID_CLIENTE " & _
                   " LEFT JOIN CLIENTES_PEDIDOS CP ON AP.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE AP.DOC_ID = 0 " & _
                   filtro & _
                   " GROUP BY AL.ID_LOTE, AP.CLIENTE_ID,AP.PEDIDO_ID " & _
                   " ORDER BY AL.ID_LOTE ASC,AP.CLIENTE_ID ASC"
        Set Listado_Pendientes_Facturar = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_LOTE,ALODINE_ID FROM ALODINE_LOTES ORDER BY ALODINE_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
'Public Function Proximo_NUMERO_LOTE(producto As String, LOTE As String, ANNO As Integer) As String
Public Function Proximo_NUMERO_LOTE(id_alodine As Long, ANNO As Integer) As String
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(NUMERO_LOTE) " & _
                   "  FROM ALODINE_LOTES A, ALODINE B " & _
                   " WHERE A.ALODINE_ID = B.ID_ALODINE " & _
                   "   AND B.ID_ALODINE = " & id_alodine & _
                   "   AND YEAR(FECHA_ALTA) = " & ANNO
'                   "   AND B.PRODUCTO = '" & producto & "'"
'                   "   AND B.LOTE = '" & LOTE & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                Proximo_NUMERO_LOTE = 1
        Else
                Proximo_NUMERO_LOTE = rs.Fields(0) + 1
        End If
        Proximo_NUMERO_LOTE = Proximo_NUMERO_LOTE
        Set rs = Nothing
End Function
Public Function recuperarMuestraEC(ALODINE_ID As Long) As Long
    Dim idMuestra As Long
    Dim consulta As String
    Dim rs As ADODB.Recordset
    idMuestra = 0
    Dim oalodine As New clsAlodine
    oalodine.Carga ALODINE_ID
    Dim listaMuestras As String
    listaMuestras = oalodine.getMUESTRAS
    If listaMuestras <> "" Then
        consulta = "select a.id_muestra from muestras a,tipos_muestra b " & _
                   " where id_muestra in (" & listaMuestras & ") " & _
                   "   and a.TIPO_MUESTRA_ID = b.ID_TIPO_MUESTRA " & _
                   "   and a.TIPO_MUESTRA_ID = 40 " & _
                   " order by id_muestra desc "
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            If Not IsNull(rs(0)) Then
                idMuestra = rs(0)
            End If
        End If
        Set rs = Nothing
    End If
    recuperarMuestraEC = idMuestra
End Function
