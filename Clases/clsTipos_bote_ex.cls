VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_bote_ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_BOTE_EX As Integer
Private TIPO_REACTIVO_EX_ID As Integer
Private TIPO_M_REFERENCIA_ID As Integer
Private CODIGO_PROVEEDOR As String
Private RESTRICCIONES As String
Private CODIGO As String
Private Calidad As String
Private PROVEEDOR_ID As Integer
Private cantidad As String
Private PRECIO As String
Private TAMANO_ETIQUETA_ID As Integer
Private CANTIDAD_MINIMA_PEDIDO As Integer
Private STOCK_MINIMO As Integer
Private TIPO_CADUCIDAD_ID As Integer
Private TEMPERATURA As String
Private HUMEDAD As String
Private otros As String
Private CANTIDAD_UNIDAD_PEDIDO As Long
Private RESPONSABLE_ID As Long
Private CODIGO_IDEN_LOTE As String
Private TEXTO_CERTIFICADO As String
Private anulado As Integer

Public Property Let setID_TIPO_BOTE_EX(ByVal dato As Integer)
    ID_TIPO_BOTE_EX = dato
End Property
Public Property Let setTIPO_REACTIVO_EX_ID(ByVal dato As Integer)
    TIPO_REACTIVO_EX_ID = dato
End Property
Public Property Let setTIPO_M_REFERENCIA_ID(ByVal dato As Integer)
    TIPO_M_REFERENCIA_ID = dato
End Property
Public Property Let setTAMANO_ETIQUETA_ID(ByVal dato As Integer)
    TAMANO_ETIQUETA_ID = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Integer)
    PROVEEDOR_ID = dato
End Property
Public Property Let setCODIGO_PROVEEDOR(ByVal dato As String)
    CODIGO_PROVEEDOR = Trim(dato)
End Property
Public Property Let setRESTRICCIONES(ByVal dato As String)
    RESTRICCIONES = Trim(dato)
End Property
Public Property Let setCODIGO(ByVal dato As String)
    CODIGO = Trim(dato)
End Property
Public Property Let setCALIDAD(ByVal dato As String)
    Calidad = Trim(dato)
End Property
Public Property Let setCANTIDAD(ByVal dato As String)
    cantidad = Trim(dato)
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = Trim(dato)
End Property
Public Property Let setCANTIDAD_MINIMA_PEDIDO(ByVal dato As Integer)
    CANTIDAD_MINIMA_PEDIDO = dato
End Property
Public Property Let setSTOCK_MINIMO(ByVal dato As Integer)
    STOCK_MINIMO = dato
End Property
Public Property Let setTIPO_CADUCIDAD_ID(ByVal dato As Integer)
    TIPO_CADUCIDAD_ID = dato
End Property
Public Property Let setTEMPERATURA(ByVal dato As String)
    TEMPERATURA = Trim(dato)
End Property
Public Property Let setHUMEDAD(ByVal dato As String)
    HUMEDAD = Trim(dato)
End Property
Public Property Let setOTROS(ByVal dato As String)
    otros = Trim(dato)
End Property
Public Property Let setCANTIDAD_UNIDAD_PEDIDO(ByVal dato As Long)
    CANTIDAD_UNIDAD_PEDIDO = Trim(dato)
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
    RESPONSABLE_ID = Trim(dato)
End Property
Public Property Let setCODIGO_IDEN_LOTE(ByVal dato As String)
    CODIGO_IDEN_LOTE = Trim(dato)
End Property
Public Property Let setTEXTO_CERTIFICADO(ByVal dato As String)
    TEXTO_CERTIFICADO = Trim(dato)
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property

'**************************************************************
'**************************************************************
Public Property Get getID_TIPO_BOTE_EX() As Integer
    getID_TIPO_BOTE_EX = ID_TIPO_BOTE_EX
End Property
Public Property Get getTIPO_REACTIVO_EX_ID() As Integer
    getTIPO_REACTIVO_EX_ID = TIPO_REACTIVO_EX_ID
End Property
Public Property Get getTIPO_M_REFERENCIA_ID() As Integer
    getTIPO_M_REFERENCIA_ID = TIPO_M_REFERENCIA_ID
End Property
Public Property Get getTAMANO_ETIQUETA_ID() As Integer
    getTAMANO_ETIQUETA_ID = TAMANO_ETIQUETA_ID
End Property
Public Property Get getPROVEEDOR_ID() As Integer
    getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getCODIGO_PROVEEDOR() As String
    getCODIGO_PROVEEDOR = CODIGO_PROVEEDOR
End Property
Public Property Get getRESTRICCIONES() As String
    getRESTRICCIONES = RESTRICCIONES
End Property
Public Property Get getCODIGO() As String
    getCODIGO = CODIGO
End Property
Public Property Get getCALIDAD() As String
    getCALIDAD = Calidad
End Property
Public Property Get getCANTIDAD() As String
    getCANTIDAD = cantidad
End Property
Public Property Get getPRECIO() As String
    getPRECIO = PRECIO
End Property
Public Property Get getCANTIDAD_MINIMA_PEDIDO() As Integer
    getCANTIDAD_MINIMA_PEDIDO = CANTIDAD_MINIMA_PEDIDO
End Property
Public Property Get getSTOCK_MINIMO() As Integer
    getSTOCK_MINIMO = STOCK_MINIMO
End Property
Public Property Get getTIPO_CADUCIDAD_ID() As Integer
    getTIPO_CADUCIDAD_ID = TIPO_CADUCIDAD_ID
End Property
Public Property Get getTEMPERATURA() As String
    getTEMPERATURA = TEMPERATURA
End Property
Public Property Get getHUMEDAD() As String
    getHUMEDAD = HUMEDAD
End Property
Public Property Get getOTROS() As String
    getOTROS = otros
End Property
Public Property Get getCANTIDAD_UNIDAD_PEDIDO() As Long
    getCANTIDAD_UNIDAD_PEDIDO = CANTIDAD_UNIDAD_PEDIDO
End Property
Public Property Get getRESPONSABLE_ID() As Long
    getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getCODIGO_IDEN_LOTE() As String
    getCODIGO_IDEN_LOTE = CODIGO_IDEN_LOTE
End Property

Public Property Get getTEXTO_CERTIFICADO() As String
    getTEXTO_CERTIFICADO = TEXTO_CERTIFICADO
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Function Listado(nombre As String, CODIGO As String, tipo As String, proveedor As Long, TIPO_REACTIVO As Long, ANULADOS As Boolean, responsable As Long) As ADODB.Recordset
    Dim consulta As String
    Dim s As String
    If tipo <> "" And tipo <> "0" Then
        s = " AND TB.TIPO_M_REFERENCIA_ID = " & CInt(tipo)
    End If
    If proveedor > 0 Then
        s = s & " AND TB.PROVEEDOR_ID = " & proveedor
    End If
    If TIPO_REACTIVO > 0 Then
        s = s & " AND TB.TIPO_REACTIVO_EX_ID = " & TIPO_REACTIVO
    End If
    If responsable <> 0 Then
        s = s & " AND TB.RESPONSABLE_ID = " & responsable
    End If
    consulta = "SELECT TB.ID_TIPO_BOTE_EX,TB.CODIGO,TR.NOMBRE,P.NOMBRE,D.DESCRIPCION,TB.PRECIO " & _
               "  FROM TIPOS_BOTE_EX AS TB, PROVEEDORES AS P, TIPOS_REACTIVO_EX AS TR,decodificadora D " & _
               " WHERE TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX " & _
               "   AND TB.PROVEEDOR_ID = P.ID_PROVEEDOR " & _
               "   AND TR.NOMBRE LIKE '%" & nombre & "%' " & _
               "   AND TB.CODIGO LIKE '%" & CODIGO & "%' " & _
               "   AND TB.TIPO_M_REFERENCIA_ID = D.VALOR " & _
               "   AND D.CODIGO = " & DECODIFICADORA.REX_TIPOS & _
               s & _
               IIf(ANULADOS, " AND TB.ANULADO <> 0 ", " AND TB.ANULADO = 0") & _
               " order by TB.ID_TIPO_BOTE_EX"
    Set Listado = datos_bd(consulta)
End Function
Public Sub Imprimir_Listado(ByVal prm_sustancia As String, ByVal prm_codigo As String, ByVal prm_tipo As Long, ByVal prm_proveedor As Long, ByVal prm_anulados As Boolean)

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    Dim strCriterio As String
                
                
    strCriterio = "{decodificadora.CODIGO} = 30"
    
    If prm_tipo > 0 Then strCriterio = strCriterio & " AND {tipos_bote_ex.TIPO_M_REFERENCIA_ID} = " & prm_tipo
    If prm_proveedor > 0 Then strCriterio = strCriterio & " AND {tipos_bote_ex.PROVEEDOR_ID} = " & prm_proveedor
    If prm_sustancia <> "" Then strCriterio = strCriterio & " AND {tipos_reactivo_ex.NOMBRE} like " & Chr(34) & "*" & prm_sustancia & "*" & Chr(34)
    If prm_codigo <> "" Then strCriterio = strCriterio & " AND {tipos_bote_ex.CODIGO} like " & Chr(34) & "*" & prm_codigo & "*" & Chr(34)
    strCriterio = strCriterio & " AND {tipos_bote_ex.ANULADO} " & IIf(prm_anulados, "<>", "=") & " 0"
                
    ReDim arrNom(1)
    ReDim arrVal(1)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        
    With objRpt
        .iniciar
        .informe = "General\rptListadoBotesEx"
        .criterio = strCriterio
        .imprimir = False
        
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal

        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objRpt
    Set objRpt = Nothing

End Sub



Public Function Listado_ID(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT TB.ID_TIPO_BOTE_EX,TB.CODIGO,TR.NOMBRE,P.NOMBRE,D.DESCRIPCION,TB.PRECIO " & _
               "  FROM TIPOS_BOTE_EX AS TB, PROVEEDORES AS P, TIPOS_REACTIVO_EX AS TR,decodificadora D " & _
               " WHERE TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX " & _
               "   AND TB.PROVEEDOR_ID = P.ID_PROVEEDOR " & _
               "   AND TB.TIPO_M_REFERENCIA_ID = D.VALOR " & _
               "   AND D.CODIGO = " & DECODIFICADORA.REX_TIPOS & _
               "   AND TB.ID_TIPO_BOTE_EX = " & ID & _
               " order by TB.ID_TIPO_BOTE_EX"
    Set Listado_ID = datos_bd(consulta)
End Function
Public Function Listado_Reactivos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT B.ID_TIPO_REACTIVO_EX, B.NOMBRE FROM TIPOS_BOTE_EX A, TIPOS_REACTIVO_EX B " & _
               " WHERE B.ID_TIPO_REACTIVO_EX = A.TIPO_REACTIVO_EX_ID " & _
               " ORDER BY B.NOMBRE ASC"
    Set Listado_Reactivos = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT ID_TIPO_BOTE_EX, CODIGO " & _
               "  FROM TIPOS_BOTE_EX " & _
               " ORDER BY CODIGO ASC"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function CARGAR(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_BOTE_EX WHERE ID_TIPO_BOTE_EX=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID_TIPO_BOTE_EX = rs("ID_TIPO_BOTE_EX")
        TIPO_REACTIVO_EX_ID = rs("TIPO_REACTIVO_EX_ID")
        TIPO_M_REFERENCIA_ID = rs("TIPO_M_REFERENCIA_ID")
        TAMANO_ETIQUETA_ID = rs("TAMANO_ETIQUETA_ID")
        PROVEEDOR_ID = rs("PROVEEDOR_ID")
        CODIGO_PROVEEDOR = rs("CODIGO_PROVEEDOR")
        RESTRICCIONES = rs("RESTRICCIONES")
        CODIGO = rs("CODIGO")
        Calidad = rs("CALIDAD")
        cantidad = rs("CANTIDAD")
        PRECIO = rs("PRECIO")
        CANTIDAD_UNIDAD_PEDIDO = rs("CANTIDAD_UNIDAD_PEDIDO")
        RESPONSABLE_ID = rs("RESPONSABLE_ID")
        CODIGO_IDEN_LOTE = rs("CODIGO_IDEN_LOTE")
        CANTIDAD_MINIMA_PEDIDO = rs("CANTIDAD_MINIMA_PEDIDO")
        STOCK_MINIMO = rs("STOCK_MINIMO")
        TIPO_CADUCIDAD_ID = rs("TIPO_CADUCIDAD_ID")
        TEMPERATURA = rs("TEMPERATURA")
        HUMEDAD = rs("HUMEDAD")
        otros = rs("OTROS")
        TEXTO_CERTIFICADO = rs("TEXTO_CERTIFICADO")
        anulado = rs("ANULADO")
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar el bote de reactivo.", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_BOTE_EX) FROM TIPOS_BOTE_EX"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_BOTE_EX = 1
    Else
        ID_TIPO_BOTE_EX = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    If duplicado = False Then
        Me.CrearID
        
        consulta = "Insert into TIPOS_BOTE_EX " & _
                   " values(" & ID_TIPO_BOTE_EX & "," & TIPO_REACTIVO_EX_ID & "," & _
                   TIPO_M_REFERENCIA_ID & ",'" & CODIGO_PROVEEDOR & "','" & RESTRICCIONES & "','" & _
                   CODIGO & "','" & Calidad & "'," & PROVEEDOR_ID & ",'" & cantidad & "','" & _
                   PRECIO & "'," & TAMANO_ETIQUETA_ID & "," & _
                   CANTIDAD_MINIMA_PEDIDO & "," & STOCK_MINIMO & "," & _
                   TIPO_CADUCIDAD_ID & ",'" & TEMPERATURA & "','" & _
                   HUMEDAD & "','" & otros & "'," & CANTIDAD_UNIDAD_PEDIDO & "," & RESPONSABLE_ID & ",'" & _
                   CODIGO_IDEN_LOTE & "','" & TEXTO_CERTIFICADO & "'," & anulado & ")"
        execute_bd consulta
        Insertar = ID_TIPO_BOTE_EX
    Else
        MsgBox "El codigo del reactivo ya existe.", vbInformation, App.Title
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Reactivo (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
'    Dim rs As New ADODB.RecordSet
'    Dim consulta As String
'    Set rs = datos_bd("SELECT ID_BOTE_EX FROM BOTES_EX WHERE TIPO_BOTE_EX_ID = " & ID)
'    If rs.RecordCount > 0 Then
'        MsgBox "No puede eliminarse el Bote. Esta usado en el reactivo número : " & rs(0)
'        Eliminar = False
'    Else
'        consulta = "DELETE FROM TIPOS_BOTE_EX WHERE ID_TIPO_BOTE_EX=" & ID
        execute_bd "UPDATE TIPOS_BOTE_EX SET ANULADO=" & USUARIO.getID_EMPLEADO & " WHERE ID_TIPO_BOTE_EX=" & ID
        Eliminar = True
'    End If
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el Bote de Reactivo (Eliminar)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE TIPOS_BOTE_EX SET " & _
                " TIPO_REACTIVO_EX_ID = " & TIPO_REACTIVO_EX_ID & "," & _
                " TIPO_M_REFERENCIA_ID = " & TIPO_M_REFERENCIA_ID & "," & _
                " TAMANO_ETIQUETA_ID = " & TAMANO_ETIQUETA_ID & "," & _
                " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                " CODIGO_PROVEEDOR = '" & CODIGO_PROVEEDOR & "'," & _
                " RESTRICCIONES = '" & RESTRICCIONES & "'," & _
                " CODIGO = '" & CODIGO & "'," & _
                " CALIDAD = '" & Calidad & "'," & _
                " CANTIDAD = '" & cantidad & "'," & _
                " CANTIDAD_MINIMA_PEDIDO = " & CANTIDAD_MINIMA_PEDIDO & "," & _
                " STOCK_MINIMO = " & STOCK_MINIMO & "," & _
                " TIPO_CADUCIDAD_ID =" & TIPO_CADUCIDAD_ID & "," & _
                " CANTIDAD_UNIDAD_PEDIDO = " & CANTIDAD_UNIDAD_PEDIDO & "," & _
                " RESPONSABLE_ID = " & RESPONSABLE_ID & "," & _
                " CODIGO_IDEN_LOTE = '" & CODIGO_IDEN_LOTE & "', " & _
                " TEMPERATURA = '" & TEMPERATURA & "'," & _
                " HUMEDAD = '" & HUMEDAD & "'," & _
                " OTROS = '" & otros & "'," & _
                " TEXTO_CERTIFICADO = '" & TEXTO_CERTIFICADO & "'," & _
                " PRECIO = '" & PRECIO & "'" & _
           " WHERE ID_TIPO_BOTE_EX = " & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar el Reactivo (Modificar)", vbCritical, Err.Description
End Function

Public Function duplicado() As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    ' No verificar duplicado
    duplicado = False
    ' Verificar duplicado
    consulta = "SELECT ID_TIPO_BOTE_EX FROM TIPOS_BOTE_EX WHERE CODIGO = '" & CODIGO & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        duplicado = False
    Else
        ID_TIPO_BOTE_EX = rs(0)
        duplicado = True
    End If
    Exit Function
fallo:
    duplicado = False
End Function
Public Function Listado_Por_Proveedor(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT TB.ID_TIPO_BOTE_EX AS ID,Concat('(',TB.CODIGO,') ',TR.NOMBRE) AS DES" & _
               "  FROM TIPOS_BOTE_EX AS TB, TIPOS_REACTIVO_EX AS TR" & _
               " WHERE TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX " & _
               "   AND TB.PROVEEDOR_ID = " & ID & _
               " ORDER BY TB.ID_TIPO_BOTE_EX"
    Set Listado_Por_Proveedor = datos_bd(consulta)
End Function
Public Function Stock(TIPO_BOTE As Long) As Integer
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo Stock_Error

        consulta = "SELECT COUNT(*) " & _
                   "  FROM BOTES_EX " & _
                   " WHERE TIPO_BOTE_EX_ID = " & TIPO_BOTE & _
                   "   AND ANULADO = 0 AND FECHA_FIN = '0000-00-00'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Stock = rs(0)
        Else
            Stock = 0
        End If

   On Error GoTo 0
   Exit Function

Stock_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Stock of Módulo de clase clsTipos_bote_ex"
End Function
'M1144-I
Public Function Recuperar(ID As Long) As Boolean
    On Error GoTo fallo
        execute_bd "UPDATE TIPOS_BOTE_EX SET ANULADO=0 WHERE ID_TIPO_BOTE_EX=" & ID
        Recuperar = True
    Exit Function
fallo:
    Recuperar = False
    MsgBox "Error al recuperar el Bote de Reactivo (Recuperar)", vbCritical, Err.Description
End Function
'M1144-F
