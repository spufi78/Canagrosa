VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_tipos_ensayos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_TIPOS_ENSAYOS
'***************************************************************
Private ID_TIPO_ENSAYO As Long
Private TIPO_ANALISIS_ID As Long
Private nombre As String
Private PROCESO_BASE_ID As Long
Private criterio As String
Private INCLUYE_ESPESOR As Integer
Private LOTE_PROBETAS As Integer
Private SEGUN_NORMA As String
Private RANGO_MIN As String
Private RANGO_MAX As String
Private RANGO_MIN_TA As String
Private RANGO_MAX_TA As String
Private RANGO_MIN_HUMEDAD As String
Private RANGO_MAX_HUMEDAD As String
Private horas As String
Private NORMA As String
Private NORMA_ID_CA As Long
Private NORMA_ID_ENSAYO As Long
Private INCLUYE_EQUIPO As Integer
Private duplicado As Integer
Private ACTIVO As Integer
Private ENAC As Integer
Private nadcap As Integer
Private PNT_VINCULADO As Integer
Private CONDICIONES_AMBIENTALES As Integer
Private OBSERVACIONES As String
Private ESPINTURA As Integer
Private ESSIMPLIFICADO As Integer
Private FORMULA_ID As Integer
Private UNIDAD_ID As Integer
Private ES_SUBCONTRATABLE As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_TIPOS_ENSAYOS
'***************************************************************
Public Property Let setID_TIPO_ENSAYO(ByVal dato As Long)
        ID_TIPO_ENSAYO = dato
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Long)
        TIPO_ANALISIS_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setPROCESO_BASE_ID(ByVal dato As Long)
        PROCESO_BASE_ID = dato
End Property
Public Property Let setCRITERIO(ByVal dato As String)
        criterio = dato
End Property
Public Property Let setINCLUYE_ESPESOR(ByVal dato As Integer)
        INCLUYE_ESPESOR = dato
End Property
Public Property Let setLOTE_PROBETAS(ByVal dato As Integer)
        LOTE_PROBETAS = dato
End Property
Public Property Let setSEGUN_NORMA(ByVal dato As String)
        SEGUN_NORMA = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As String)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As String)
        RANGO_MAX = dato
End Property
Public Property Let setRANGO_MIN_TA(ByVal dato As String)
        RANGO_MIN_TA = dato
End Property
Public Property Let setRANGO_MAX_TA(ByVal dato As String)
        RANGO_MAX_TA = dato
End Property
Public Property Let setRANGO_MIN_HUMEDAD(ByVal dato As String)
        RANGO_MIN_HUMEDAD = dato
End Property
Public Property Let setRANGO_MAX_HUMEDAD(ByVal dato As String)
        RANGO_MAX_HUMEDAD = dato
End Property
Public Property Let setHORAS(ByVal dato As String)
        horas = dato
End Property
Public Property Let setNORMA(ByVal dato As String)
        NORMA = dato
End Property
Public Property Let setNORMA_ID_CA(ByVal dato As Long)
        NORMA_ID_CA = dato
End Property
Public Property Let setNORMA_ID_ENSAYO(ByVal dato As Long)
        NORMA_ID_ENSAYO = dato
End Property
Public Property Let setINCLUYE_EQUIPO(ByVal dato As Integer)
        INCLUYE_EQUIPO = dato
End Property
Public Property Let setDUPLICADO(ByVal dato As Integer)
        duplicado = dato
End Property
Public Property Let setACTIVO(ByVal dato As Integer)
        ACTIVO = dato
End Property
Public Property Let setENAC(ByVal dato As Integer)
        ENAC = dato
End Property
Public Property Let setNADCAP(ByVal dato As Integer)
        nadcap = dato
End Property
Public Property Let setPNT_VINCULADO(ByVal dato As Integer)
        PNT_VINCULADO = dato
End Property
Public Property Let setCONDICIONES_AMBIENTALES(ByVal dato As Integer)
        CONDICIONES_AMBIENTALES = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setESPINTURA(ByVal dato As Integer)
        ESPINTURA = dato
End Property
Public Property Let setESSIMPLIFICADO(ByVal dato As Integer)
        ESSIMPLIFICADO = dato
End Property
Public Property Let setFORMULA_ID(ByVal dato As Integer)
        FORMULA_ID = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Integer)
        UNIDAD_ID = dato
End Property
'MXXXX-I
Public Property Let setES_SUBCONTRATABLE(ByVal dato As Integer)
        ES_SUBCONTRATABLE = dato
End Property
'MXXXX-F
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_TIPOS_ENSAYOS
'***************************************************************
Public Property Get getID_TIPO_ENSAYO() As Long
        getID_TIPO_ENSAYO = ID_TIPO_ENSAYO
End Property
Public Property Get getTIPO_ANALISIS_ID() As Long
    getTIPO_ANALISIS_ID = TIPO_ANALISIS_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getPROCESO_BASE_ID() As Long
        getPROCESO_BASE_ID = PROCESO_BASE_ID
End Property
Public Property Get getCRITERIO() As String
        getCRITERIO = criterio
End Property
Public Property Get getINCLUYE_ESPESOR() As Integer
        getINCLUYE_ESPESOR = INCLUYE_ESPESOR
End Property
Public Property Get getLOTE_PROBETAS() As Integer
        getLOTE_PROBETAS = LOTE_PROBETAS
End Property
Public Property Get getSEGUN_NORMA() As String
        getSEGUN_NORMA = SEGUN_NORMA
End Property
Public Property Get getRANGO_MIN() As String
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As String
        getRANGO_MAX = RANGO_MAX
End Property
Public Property Get getRANGO_MIN_TA() As String
        getRANGO_MIN_TA = RANGO_MIN_TA
End Property
Public Property Get getRANGO_MAX_TA() As String
        getRANGO_MAX_TA = RANGO_MAX_TA
End Property
Public Property Get getRANGO_MIN_HUMEDAD() As String
        getRANGO_MIN_HUMEDAD = RANGO_MIN_HUMEDAD
End Property
Public Property Get getRANGO_MAX_HUMEDAD() As String
        getRANGO_MAX_HUMEDAD = RANGO_MAX_HUMEDAD
End Property
Public Property Get getHORAS() As String
        getHORAS = horas
End Property
Public Property Get getNORMA() As String
        getNORMA = NORMA
End Property
Public Property Get getNORMA_ID_CA() As Long
        getNORMA_ID_CA = NORMA_ID_CA
End Property
Public Property Get getNORMA_ID_ENSAYO() As Long
        getNORMA_ID_ENSAYO = NORMA_ID_ENSAYO
End Property
Public Property Get getINCLUYE_EQUIPO() As Integer
        getINCLUYE_EQUIPO = INCLUYE_EQUIPO
End Property
Public Property Get getDUPLICADO() As Integer
        getDUPLICADO = duplicado
End Property
Public Property Get getACTIVO() As Integer
        getACTIVO = ACTIVO
End Property
Public Property Get getENAC() As Integer
        getENAC = ENAC
End Property
Public Property Get getNADCAP() As Integer
        getNADCAP = nadcap
End Property
Public Property Get getPNT_VINCULADO() As Integer
        getPNT_VINCULADO = PNT_VINCULADO
End Property
Public Property Get getCONDICIONES_AMBIENTALES() As Integer
        getCONDICIONES_AMBIENTALES = CONDICIONES_AMBIENTALES
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getESPINTURA() As Integer
        getESPINTURA = ESPINTURA
End Property
Public Property Get getESSIMPLIFICADO() As Integer
        getESSIMPLIFICADO = ESSIMPLIFICADO
End Property
Public Property Get getFORMULA_ID() As Integer
        getFORMULA_ID = FORMULA_ID
End Property
Public Property Get getUNIDAD_ID() As Integer
        getUNIDAD_ID = UNIDAD_ID
End Property
'MXXXX-I
Public Property Get getES_SUBCONTRATABLE() As Integer
        getES_SUBCONTRATABLE = ES_SUBCONTRATABLE
End Property
'MXXXX-F
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_TIPOS_ENSAYOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_TIPOS_ENSAYOS WHERE ID_TIPO_ENSAYO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
                MsgBox "Error al cargar los datos del control de eficacia.", vbCritical, App.Title
        Else
                ID_TIPO_ENSAYO = rs("ID_TIPO_ENSAYO")
                TIPO_ANALISIS_ID = rs("TIPO_ANALISIS_ID")
                nombre = rs("NOMBRE")
                PROCESO_BASE_ID = rs("PROCESO_BASE_ID")
                criterio = rs("CRITERIO")
                INCLUYE_ESPESOR = rs("INCLUYE_ESPESOR")
                LOTE_PROBETAS = rs("LOTE_PROBETAS")
                SEGUN_NORMA = rs("SEGUN_NORMA")
                RANGO_MIN = rs("RANGO_MIN")
                RANGO_MAX = rs("RANGO_MAX")
                RANGO_MIN_TA = IIf(IsNull(rs("RANGO_MIN_TA")), "", rs("RANGO_MIN_TA"))
                RANGO_MAX_TA = IIf(IsNull(rs("RANGO_MAX_TA")), "", rs("RANGO_MAX_TA"))
                RANGO_MIN_HUMEDAD = IIf(IsNull(rs("RANGO_MIN_HUMEDAD")), "", rs("RANGO_MIN_HUMEDAD"))
                RANGO_MAX_HUMEDAD = IIf(IsNull(rs("RANGO_MAX_HUMEDAD")), "", rs("RANGO_MAX_HUMEDAD"))
                horas = rs("HORAS")
                NORMA = rs("NORMA")
                NORMA_ID_CA = rs("NORMA_ID_CA")
                NORMA_ID_ENSAYO = rs("NORMA_ID_ENSAYO")
                INCLUYE_EQUIPO = rs("INCLUYE_EQUIPO")
                duplicado = rs("DUPLICADO")
                ACTIVO = rs("ACTIVO")
                ENAC = rs("ENAC")
                nadcap = rs("NADCAP")
                PNT_VINCULADO = rs("PNT_VINCULADO")
                CONDICIONES_AMBIENTALES = rs("CONDICIONES_AMBIENTALES")
                OBSERVACIONES = rs("OBSERVACIONES")
                FORMULA_ID = rs("FORMULA_ID")
                ESPINTURA = rs("ESPINTURA")
                ESSIMPLIFICADO = rs("ESSIMPLIFICADO")
                UNIDAD_ID = rs("UNIDAD_ID")
                'MXXXX-I
                ES_SUBCONTRATABLE = rs("ES_SUBCONTRATABLE")
                'MXXXX-F
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        log ("Error al cargar los datos(clsCe_tipos_ensayos)")
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_TIPO_ENSAYO) FROM CE_TIPOS_ENSAYOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_TIPO_ENSAYO = 1
        Else
                ID_TIPO_ENSAYO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CE_TIPOS_ENSAYOS " & _
                   "  VALUES (" & _
                        ID_TIPO_ENSAYO & "," & TIPO_ANALISIS_ID & "," & "'" & nombre & "'" & "," & PROCESO_BASE_ID & "," & _
                        "'" & criterio & "'" & "," & INCLUYE_ESPESOR & "," & LOTE_PROBETAS & "," & _
                        "'" & SEGUN_NORMA & "','" & RANGO_MIN & "','" & RANGO_MAX & "'," & _
                        numerico_bd(RANGO_MIN_TA) & "," & numerico_bd(RANGO_MAX_TA) & "," & _
                        numerico_bd(RANGO_MIN_HUMEDAD) & "," & numerico_bd(RANGO_MAX_HUMEDAD) & "," & _
                        "'" & horas & "'," & _
                        "'" & NORMA & "'" & "," & NORMA_ID_CA & "," & NORMA_ID_ENSAYO & "," & INCLUYE_EQUIPO & "," & duplicado & "," & ACTIVO & "," & ENAC & "," & nadcap & "," & _
                        PNT_VINCULADO & ",'" & CONDICIONES_AMBIENTALES & "','" & OBSERVACIONES & "'," & ESPINTURA & "," & ESSIMPLIFICADO & "," & FORMULA_ID & "," & UNIDAD_ID & "," & ES_SUBCONTRATABLE & _
                ")"
        execute_bd consulta
        Insertar = ID_TIPO_ENSAYO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_tipos_ensayos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_TIPOS_ENSAYOS SET " & _
                        " TIPO_ANALISIS_ID = " & TIPO_ANALISIS_ID & "," & _
                        " NOMBRE = '" & nombre & "'," & _
                        " PROCESO_BASE_ID = " & PROCESO_BASE_ID & "," & _
                        " CRITERIO = '" & criterio & "'," & _
                        " INCLUYE_ESPESOR = " & INCLUYE_ESPESOR & "," & _
                        " LOTE_PROBETAS = " & LOTE_PROBETAS & "," & _
                        " SEGUN_NORMA = '" & SEGUN_NORMA & "'," & _
                        " RANGO_MIN = '" & RANGO_MIN & "'," & " RANGO_MAX = '" & RANGO_MAX & "'," & _
                        " RANGO_MIN_TA = " & RANGO_MIN_TA & "," & " RANGO_MAX_TA = " & RANGO_MAX_TA & "," & _
                        " RANGO_MIN_HUMEDAD = " & RANGO_MIN_HUMEDAD & "," & " RANGO_MAX_HUMEDAD = " & RANGO_MAX_HUMEDAD & "," & _
                        " HORAS = '" & horas & "'," & _
                        " NORMA = '" & NORMA & "'," & _
                        " NORMA_ID_CA = " & NORMA_ID_CA & "," & " NORMA_ID_ENSAYO = " & NORMA_ID_ENSAYO & "," & _
                        " INCLUYE_EQUIPO = " & INCLUYE_EQUIPO & "," & _
                        " DUPLICADO = " & duplicado & "," & " ACTIVO = " & ACTIVO & "," & " ENAC = " & ENAC & ", NADCAP = " & nadcap & "," & _
                        " PNT_VINCULADO = " & PNT_VINCULADO & "," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "," & " CONDICIONES_AMBIENTALES = '" & CONDICIONES_AMBIENTALES & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "', ESPINTURA = " & ESPINTURA & ",ESSIMPLIFICADO=" & ESSIMPLIFICADO & "," & _
                        " FORMULA_ID = " & FORMULA_ID & "," & _
                        " ES_SUBCONTRATABLE = " & ES_SUBCONTRATABLE & _
                " WHERE ID_TIPO_ENSAYO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_tipos_ensayos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        'M1103-I
        Dim rs As ADODB.Recordset
        consulta = "SELECT MUESTRA_ID FROM CE_RECEPCION WHERE TIPO_ENSAYO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar el ENSAYO. Existen muestras asociadas. Si lo desea, puede DESACTIVARLO desde el Detalle.", vbExclamation, App.Title
            Eliminar = False
            Exit Function
        End If
        'M1103-F
        execute_bd ("DELETE FROM CE_TIPOS_ENSAYOS WHERE ID_TIPO_ENSAYO = " & ID)
        execute_bd ("DELETE FROM CE_TIPOS_ENSAYOS_EQUIPOS WHERE TIPO_ENSAYO_ID = " & ID)
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_tipos_ensayos)", vbCritical, Err.Description
End Function
Public Function Listado(nombre As String, PROCESO As String, lACTIVA As Boolean, lSubcontrata As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim s As String, s2 As String, s3 As String
        If PROCESO <> "" Then
            s = "   AND B.NOMBRE LIKE '%" & PROCESO & "%'"
        End If
        If lACTIVA = True Then
            s2 = " AND A.ACTIVO = 1 "
        End If
'M1147-I
        If lSubcontrata = True Then
            s3 = " AND A.ES_SUBCONTRATABLE = 1 "
        End If
'M1147-F
        consulta = "SELECT A.NOMBRE,B.NOMBRE,A.ID_TIPO_ENSAYO,A.NORMA,A.ES_SUBCONTRATABLE " & _
                   "  FROM CE_TIPOS_ENSAYOS A LEFT JOIN PROCESOS_BASE B " & _
                   "    ON A.PROCESO_BASE_ID = B.ID_PROCESO_BASE " & _
                   " WHERE A.NOMBRE LIKE '%" & nombre & "%'" & _
                   s & s2 & s3 & _
                   " ORDER BY A.NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_TIPO_ENSAYO,CONCAT(NOMBRE,' (Id:',ID_TIPO_ENSAYO,')') FROM CE_TIPOS_ENSAYOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CE_TIPOS_ENSAYOS"
        .setDESCRIPCION = "Tipos Ensayos Eficacia"
        .setPK = "ID_TIPO_ENSAYO"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        .setFILTRO = filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Equipos_Listado(TIPO_ENSAYO_ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.EQUIPO_ID,B.NOMBRE,B.SERIE,A.EN_INFORME,A.ORDEN " & _
                   "  FROM CE_TIPOS_ENSAYOS_EQUIPOS A, EQUIPOS B" & _
                   " WHERE A.TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID & _
                   "   AND A.EQUIPO_ID = B.ID_EQUIPO " & _
                   " ORDER BY A.ORDEN"
        Set Equipos_Listado = datos_bd(consulta)
End Function
'M1137-I
Public Function Equipos_Listado_Recepcion(TIPO_ENSAYO_ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.EQUIPO_ID,B.NOMBRE,B.SERIE,A.EN_INFORME,A.ORDEN " & _
                   "  FROM CE_TIPOS_ENSAYOS_EQUIPOS A, EQUIPOS B" & _
                   " WHERE A.TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID & _
                   "   AND A.EQUIPO_ID = B.ID_EQUIPO " & _
                   "   AND B.ESTADO_ID IN ('" & ENUM_EQ_ESTADOS_ACTIVO & "','" & ENUM_EQ_ESTADOS_NRC & "')" & _
                   " ORDER BY A.ORDEN"
        Set Equipos_Listado_Recepcion = datos_bd(consulta)
End Function
'M1137-F

Public Function Equipos_Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd ("DELETE FROM CE_TIPOS_ENSAYOS_EQUIPOS WHERE TIPO_ENSAYO_ID = " & ID)
        Equipos_Eliminar = True
        Exit Function
fallo:
        Equipos_Eliminar = False
        MsgBox "Error al Equipos_Eliminar (clsCe_tipos_ensayos)", vbCritical, Err.Description
End Function
Public Function Equipos_Insertar(tipo_ensayo As Long, EQUIPO As Long, ORDEN As Integer, informe As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "INSERT INTO CE_TIPOS_ENSAYOS_EQUIPOS VALUES (" & tipo_ensayo & "," & EQUIPO & "," & ORDEN & "," & Abs(informe) & ")"
        Equipos_Insertar = True
        Exit Function
fallo:
        Equipos_Insertar = False
        MsgBox "Error al Equipos_Insertar (clsCe_tipos_ensayos)", vbCritical, Err.Description
End Function

Public Function duplicar(ID_TIPO_ENSAYO As Long) As Long
    Dim oCE_TE As New clsCe_tipos_ensayos
    Dim ote As Long
   On Error GoTo duplicar_Error

    ote = 0
    If oCE_TE.Carga(ID_TIPO_ENSAYO) Then
        ' Tipos de Ensayos
        With oCE_TE
            .setNOMBRE = oCE_TE.getNOMBRE & " (Duplicado)"
            ote = .Insertar
            If ote = 0 Then
                MsgBox "Error al duplicar el tipo de ensayo de eficacia.", vbCritical, App.Title
                duplicar = ote
                Exit Function
            End If
        End With
        Dim c As String
        ' ce_tipos_ensayos_botes_ex
        c = "INSERT INTO ce_tipos_ensayos_botes_ex "
        c = c & " SELECT " & ote & ",BOTE_EX_ID,TIPO,ORDEN FROM ce_tipos_ensayos_botes_ex "
        c = c & " WHERE TIPO_ENSAYO_ID = " & ID_TIPO_ENSAYO
        execute_bd c
        ' ce_tipos_ensayos_equipos
        c = "INSERT INTO ce_tipos_ensayos_equipos "
        c = c & " SELECT " & ote & ",EQUIPO_ID,ORDEN,EN_INFORME FROM ce_tipos_ensayos_equipos "
        c = c & " WHERE TIPO_ENSAYO_ID = " & ID_TIPO_ENSAYO
        execute_bd c
    End If

    
    
    duplicar = ote

   On Error GoTo 0
   Exit Function

duplicar_Error:
    duplicar = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure duplicar of Módulo de clase clsCe_tipos_ensayos"
End Function

