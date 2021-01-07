VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDecodificadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Private CODIGO As Long
Private idioma As String
Private VALOR As Long
Private DESCRIPCION As String
Private parametros As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Public Property Let setCODIGO(ByVal dato As Long)
        CODIGO = dato
End Property
Public Property Let setIDIOMA(ByVal dato As String)
        idioma = dato
End Property
Public Property Let setVALOR(ByVal dato As Long)
        VALOR = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setPARAMETROS(ByVal dato As String)
        parametros = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Public Property Get getCODIGO() As Long
        getCODIGO = CODIGO
End Property
Public Property Get getIDIOMA() As String
        getIDIOMA = idioma
End Property
Public Property Get getVALOR() As Long
        getVALOR = VALOR
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPARAMETROS() As String
        getPARAMETROS = parametros
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Public Function Crear_nuevo_valor()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(VALOR) FROM decodificadora WHERE CODIGO = " & CODIGO
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                VALOR = 1
        Else
                VALOR = CInt(rs.Fields(0)) + 1
        End If
        Set rs = Nothing
End Function

Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM decodificadora WHERE CODIGO = 0 AND VALOR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CODIGO = rs("CODIGO")
                idioma = rs("IDIOMA")
                VALOR = rs("VALOR")
                DESCRIPCION = rs("DESCRIPCION")
                parametros = rs("PARAMETROS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Carga_valor(lCODIGO As Long, lVALOR As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM decodificadora WHERE CODIGO = " & lCODIGO & " AND VALOR = " & lVALOR
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_valor = False
        Else
                CODIGO = rs("CODIGO")
                idioma = rs("IDIOMA")
                VALOR = rs("VALOR")
                DESCRIPCION = rs("DESCRIPCION")
                parametros = rs("PARAMETROS")
                Carga_valor = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_valor = False
        MsgBox "Error al cargar los datos(clsDecodificadora)", vbCritical, Err.Description
End Function

Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM decodificadora WHERE CODIGO = " & CODIGO & " AND DESCRIPCION = '" & Trim(DESCRIPCION) & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            If Trim(VALOR) = 0 Then
                Me.Crear_nuevo_valor
            End If
            consulta = "INSERT INTO decodificadora " & _
                       "  VALUES (" & _
                            CODIGO & "," & "'" & idioma & "'" & "," & VALOR & "," & _
                            "'" & DESCRIPCION & "','" & parametros & "'" & _
                    ")"
            execute_bd consulta
        End If
        Insertar = CODIGO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function InsertarDuplicada() As String
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        If Trim(VALOR) = 0 Then
            Me.Crear_nuevo_valor
        End If
        consulta = "INSERT INTO decodificadora " & _
                   "  VALUES (" & _
                        CODIGO & "," & "'" & idioma & "'" & "," & VALOR & "," & _
                        "'" & DESCRIPCION & "','" & parametros & "'" & _
                ")"
        execute_bd consulta
        InsertarDuplicada = CODIGO
        Exit Function
fallo:
        InsertarDuplicada = 0
        MsgBox "Error al insertar (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Modificar(cCODIGO As Long, cVALOR As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE decodificadora SET " & _
                        " IDIOMA = '" & idioma & "'," & _
                        " PARAMETROS = '" & parametros & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE CODIGO = " & cCODIGO & " AND VALOR = " & cVALOR
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Eliminar(cCODIGO As Long, cVALOR As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        If Me.Verificar_integridad(cCODIGO, cVALOR) Then
            consulta = "DELETE FROM decodificadora " & _
                    "    WHERE CODIGO = " & cCODIGO & " AND VALOR = " & cVALOR
            execute_bd consulta
        End If
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function EliminarCodigo(cCODIGO As Long, PARAMETRO As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM decodificadora " & _
                   "    WHERE CODIGO = " & cCODIGO & " AND PARAMETROS = '" & PARAMETRO & "'"
        execute_bd consulta
        EliminarCodigo = True
        Exit Function
fallo:
        EliminarCodigo = False
        MsgBox "Error al EliminarCodigo (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Listado(CODIGO) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM decodificadora " & _
                   " WHERE CODIGO = " & CODIGO & _
                   "   AND VALOR <> 0 " & _
                   " ORDER BY DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoDescripcionProducto(CODIGO As Integer, des As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If des <> "" Then
            filtro = " and a.descripcion like '%" & des & "%'"
        End If
        consulta = "SELECT a.valor, a.descripcion, b.nombre, b.id_tipo_muestra FROM decodificadora a left join tipos_muestra b on a.parametros = b.id_tipo_muestra " & _
                   " WHERE a.CODIGO = " & CODIGO & _
                   "   AND a.VALOR <> 0 " & _
                   filtro & _
                   " ORDER BY b.nombre,a.DESCRIPCION"
        Set ListadoDescripcionProducto = datos_bd(consulta)
End Function

Public Function Listado_por_Codigo(CODIGO) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM decodificadora " & _
                   " WHERE CODIGO = " & CODIGO & _
                   "   AND VALOR <> 0 " & _
                   " ORDER BY VALOR"
        Set Listado_por_Codigo = datos_bd(consulta)
End Function
Public Function Listado_por_Codigo_Parametro(CODIGO As Integer, PARAMETRO As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM decodificadora " & _
                   " WHERE CODIGO = " & CODIGO & _
                   "   AND VALOR <> 0 " & _
                   "   AND PARAMETROS = '" & PARAMETRO & "'" & _
                   " ORDER BY VALOR"
        Set Listado_por_Codigo_Parametro = datos_bd(consulta)
End Function
Public Function cargar_combo(combo As DataCombo, CCODTABL As Long) As Boolean
    Dim consulta As String
    consulta = "SELECT VALOR,DESCRIPCION FROM decodificadora WHERE CODIGO = " & CCODTABL & " order by DESCRIPCION"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    Set combo.RowSource = rs
    combo.ListField = rs(1).Name
    combo.BoundColumn = rs(0).Name
    Set rs = Nothing
End Function
Public Function cargar_mi_combo(combo As miCombo, CCODTABL As Long) As Boolean
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With combo
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "DECODIFICADORA"
            .setDESCRIPCION = "Decodificadora"
            .setQUERY = "SELECT VALOR,DESCRIPCION FROM decodificadora WHERE CODIGO = " & CCODTABL
            .setPK = "VALOR"
            .setCAMPO = "DESCRIPCION"
            .setMUESTRA_DETALLE = False
            .setFILTRO = ""
            Set .FORMULARIO = Nothing
        End With
    End If
    Set conn = Nothing
    
End Function
Public Function cargar_mi_combo_parametro(combo As miCombo, CCODTABL As Long, PARAMETRO As String) As Boolean
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With combo
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "DECODIFICADORA"
            .setDESCRIPCION = "Decodificadora"
            .setQUERY = "SELECT VALOR,DESCRIPCION FROM decodificadora WHERE CODIGO = " & CCODTABL & " AND PARAMETROS = '" & PARAMETRO & "'"
            .setPK = "VALOR"
            .setCAMPO = "DESCRIPCION"
            .setMUESTRA_DETALLE = False
            .setFILTRO = ""
            Set .FORMULARIO = Nothing
        End With
    End If
    Set conn = Nothing
    
End Function
Public Function Cargar_ComboBox(ByRef combo As ComboBox, ByRef CCODTABL As Integer) As Boolean
    Dim consulta As String
    combo.Clear
    
    consulta = "SELECT VALOR,DESCRIPCION FROM decodificadora WHERE CODIGO = " & CCODTABL & " order by DESCRIPCION"
    Dim rs As New ADODB.Recordset
    Set rs = datos_bd(consulta)
    rs.MoveFirst
    
    While Not rs.EOF
        
        combo.AddItem rs(1).Value
        combo.ItemData(combo.ListCount - 1) = rs(0).Value
        
        rs.MoveNext
    Wend
    
End Function
Public Function Verificar_integridad(cCODIGO As Long, cVALOR As Long) As Boolean
        Dim consulta As String
        Select Case cCODIGO
         Case DECODIFICADORA.CA_NORMAS_ESTADOS
            consulta = "SELECT * FROM CA_NORMAS WHERE ESTADO_ID = " & cVALOR
         Case DECODIFICADORA.CA_NORMAS_SECTORES
            consulta = "SELECT * FROM CA_NORMAS WHERE SECTOR_ID = " & cVALOR
         Case DECODIFICADORA.CA_NORMAS_TIPOS
            consulta = "SELECT * FROM CA_NORMAS WHERE TIPO_ID = " & cVALOR
         Case DECODIFICADORA.CA_NORMAS_SUBTIPOS
            consulta = "SELECT * FROM CA_NORMAS WHERE AFECTA = " & cVALOR
         Case DECODIFICADORA.CA_DOCUMENTOS_ESTADOS
            consulta = "SELECT * FROM ca_documentos WHERE ESTADO_ID = " & cVALOR
         Case DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS
            consulta = "SELECT * FROM ca_documentos WHERE FAMILIA_ID = " & cVALOR
         Case DECODIFICADORA.CA_DOCUMENTOS_RESPONSABLES
            consulta = "SELECT * FROM ca_documentos WHERE RESPONSABLE_ID = " & cVALOR
         Case DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS
            consulta = "SELECT * FROM ca_documentos WHERE SUBFAMILIA_ID = " & cVALOR
        'E0005-I
' ahora se tira de la tabla proveedores
'         Case decodificadora.EQ_PROVEEDORES
'            consulta = "SELECT * FROM EQUIPOS WHERE PROVEEDOR_ID = " & cVALOR
         Case DECODIFICADORA.EQ_FAMILIAS
            consulta = "SELECT * FROM EQUIPOS WHERE FAMILIA_ID = " & cVALOR
         Case DECODIFICADORA.EQ_SITUACIONES
            consulta = "SELECT * FROM EQUIPOS WHERE SITUACION_ID = " & cVALOR
' ahora se tira de la tabla unidades
'         Case decodificadora.EQ_UNIDADES
'            consulta = "SELECT * FROM EQUIPOS WHERE UNIDAD_ID = " & cVALOR
         Case DECODIFICADORA.EQ_periodicidad
            ' En esos tres campos son de periodicidad => Se verifica en todos
            consulta = "SELECT * FROM EQUIPOS " & _
                       "        WHERE PERIODICIDAD_CALIBRACION_ID = " & cVALOR & _
                       "           OR PERIODICIDAD_VERIFICACION_ID = " & cVALOR & _
                       "           OR PERIODICIDAD_MANTENIMIENTO_ID = " & cVALOR
                       
         Case DECODIFICADORA.EQ_FAMILIAS_ACCIONES_PLANES_MTO
              consulta = "SELECT * FROM eq_planes_acciones WHERE FAMILIA_ACC_ID = " & cVALOR
        'E0005-F
        'E0186-I
         Case DECODIFICADORA.EP_EMPRESAS_MENSAJERIA
              consulta = "SELECT * FROM EP_PAQUETES WHERE MENSAJERIA_ID = " & cVALOR
        'E0186-F
         Case DECODIFICADORA.REX_TIPOS
              consulta = "SELECT * FROM TIPOS_BOTE_EX WHERE TIPO_M_REFERENCIA_ID = " & cVALOR
         Case DECODIFICADORA.PROCNC_AUDITORIAS
             consulta = "SELECT * FROM PROCNC WHERE AUDITORIA_ID = " & cVALOR
         Case DECODIFICADORA.SC_SUBTIPOS
             consulta = "SELECT * FROM SC_PAQUETES WHERE SUBTIPO = " & cVALOR
         Case DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS
             consulta = "SELECT * FROM PROVEEDORES_FACTURAS WHERE SUBCUENTA_PAGO = " & cVALOR
         Case DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS
             consulta = "SELECT * FROM PROVEEDORES_FACTURAS WHERE SUBCUENTA_PAGO = " & cVALOR
        End Select
        If consulta <> "" Then
            Dim rs As ADODB.Recordset
            Set rs = datos_bd(consulta)
            If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                Verificar_integridad = True
            Else
                Verificar_integridad = False
                MsgBox "No se puede eliminar el registro. Tiene datos vinculados.", vbExclamation, App.Title
            End If
            Set rs = Nothing
        Else
            Verificar_integridad = True
        End If
End Function
