VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE PROVEEDOR    *
'*****************************************************
Private ID_PROVEEDOR As Long
Private nombre As String
Private DIRECCION As String
Private COD_POSTAL As Long
Private CIF As String
Private TELEFONO As String
Private FAX As String
Private responsable As String
Private CARGO As String
Private tipo As String
Private EMAIL As String
Private OBSERVACIONES As String
Private CC As String
Private CC_RETENCION As String
Private WEB As String
Private MUNICIPIO_ID As Long
Private PROVINCIA_ID As Long
Private PAIS_ID As Long
Private FP_ID As Long
Private VENCIMIENTO_ID As Long
Private BANCO_ID As Long
Private ES_SUBCONTRATA As Long
Private anulado As Integer
Private INTRA As Integer
Private EXTRA As Integer
'MXXXX-I
Private ES_FORMADOR As Long
'MXXXX-F
'M1334-I
Private ES_CUENTA_ESPANYA As Long
Private CUENTA_BANCARIA As String
Private banco As String
'M1334-F
'M1339-I
Private EMAIL_FACTURACION As String
Private EVAL_MA As Integer
Private DIAS_ENTREGA As Integer
'M1339-F


Public Sub ImprimirListadoCompleto(nombre As String, CIF As String, TELEFONO As String, booContratas As Boolean, booanulados As Boolean)
    Dim objfrm As New frmReport
    Dim criterio As String
    criterio = " {proveedores.ID_PROVEEDOR} <> 0 "
    If Trim(nombre) <> "" Then
        criterio = criterio & " AND {proveedores.NOMBRE} LIKE '*" & nombre & "*' "
    End If
    If Trim(CIF) <> "" Then
        criterio = criterio & " AND {proveedores.CIF} LIKE '*" & CIF & "*' "
    End If
    If Trim(TELEFONO) <> "" Then
        criterio = criterio & " AND {proveedores.TELEFONO} LIKE '*" & TELEFONO & "*' "
    End If
    If booContratas Then
        criterio = criterio & " AND {proveedores.ES_SUBCONTRATA} = 1 "
    End If
    If Not booanulados Then
        criterio = criterio & " AND {proveedores.ANULADO} = 0 "
    End If
    With objfrm
        .iniciar
        .informe = "Proveedores\rptListadoProveedoresCompleto"
        .criterio = criterio
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub
Public Sub ImprimirListado()
    Dim P1() As String
    Dim P2() As String
    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FECHA"
    Dim oParametros As New clsParametros
    oParametros.Carga FECHA_REVISION_PROVEEDORES, ""
    P2(1) = oParametros.getVALOR
    Dim objfrm As New frmReport
    With objfrm
        .iniciar
        .informe = "Proveedores\rptListadoProveedores"
        .criterio = "{dservicios.CODIGO} = 60.00 and {dmetodos.CODIGO} = 61.00 and {proveedores.anulado} = 0"
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub
Public Sub ImprimirListadoLI15()
    Dim P2() As String
    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FECHA"
    Dim oParametros As New clsParametros
    oParametros.Carga FECHA_REVISION_PROVEEDORES, ""
    P2(1) = oParametros.getVALOR
    Dim objfrm As New frmReport
    With objfrm
        .iniciar
        .informe = "Proveedores\rptListadoProveedoresLI15"
        .criterio = "{dservicios.CODIGO} = 60.00 and {dmetodos.CODIGO} = 61.00 and {proveedores.ES_SUBCONTRATA} = 1 and {proveedores.anulado} = 0"
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub
Public Sub ImprimirFicha(ID As Long)
    Dim objfrm As New frmReport
    With objfrm
        .iniciar
        .informe = "Proveedores\rptProveedores_Ficha"
        .criterio = "{proveedores.ID_PROVEEDOR} = " & ID
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub

'*****************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE MUESTRA    *
'*****************************************************
Public Property Let setID_PROVEEDOR(ByVal dato As Long)
    ID_PROVEEDOR = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setDIRECCION(ByVal dato As String)
    DIRECCION = Trim(dato)
End Property

Public Property Let setCOD_POSTAL(ByVal dato As Long)
    COD_POSTAL = dato
End Property
Public Property Let setCIF(ByVal dato As String)
    CIF = Trim(dato)
End Property
Public Property Let setTELEFONO(ByVal dato As String)
    TELEFONO = Trim(dato)
End Property
Public Property Let setFAX(ByVal dato As String)
    FAX = Trim(dato)
End Property
Public Property Let setRESPONSABLE(ByVal dato As String)
    responsable = Trim(dato)
End Property
Public Property Let setCARGO(ByVal dato As String)
    CARGO = Trim(dato)
End Property
Public Property Let setTIPO(ByVal dato As String)
    tipo = Trim(dato)
End Property
Public Property Let setEMAIL(ByVal dato As String)
    EMAIL = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setCC(ByVal dato As String)
    CC = Trim(dato)
End Property
Public Property Let setCC_RETENCION(ByVal dato As String)
    CC_RETENCION = Trim(dato)
End Property
Public Property Let setWEB(ByVal dato As String)
    WEB = Trim(dato)
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
Public Property Let setINTRA(ByVal dato As Integer)
    INTRA = dato
End Property
Public Property Let setEXTRA(ByVal dato As Integer)
    EXTRA = dato
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
    MUNICIPIO_ID = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setPAIS_ID(ByVal dato As Long)
    PAIS_ID = dato
End Property
Public Property Let setFP_ID(ByVal dato As Long)
    FP_ID = dato
End Property
Public Property Let setVENCIMIENTO_ID(ByVal dato As Long)
    VENCIMIENTO_ID = dato
End Property
Public Property Let setBANCO_ID(ByVal dato As Long)
    BANCO_ID = dato
End Property
'E0200-I
Public Property Let setES_SUBCONTRATA(ByVal dato As Long)
    ES_SUBCONTRATA = dato
End Property
'E0200-F
'MXXXX-I
Public Property Let setES_FORMADOR(ByVal dato As Long)
    ES_FORMADOR = dato
End Property
'MXXXX-F
'M1334-I
Public Property Let setES_CUENTA_ESPANYA(ByVal dato As Long)
    ES_CUENTA_ESPANYA = dato
End Property
Public Property Let setCUENTA_BANCARIA(ByVal dato As String)
    CUENTA_BANCARIA = Trim(dato)
End Property
Public Property Let setBANCO(ByVal dato As String)
    banco = Trim(dato)
End Property
'M1334-F
'M1339-I
Public Property Let setEMAIL_FACTURACION(ByVal dato As String)
    EMAIL_FACTURACION = Trim(dato)
End Property
'M1339-F
Public Property Let setEVAL_MA(ByVal dato As Integer)
    EVAL_MA = Trim(dato)
End Property
Public Property Let setDIAS_ENTREGA(ByVal dato As Integer)
    DIAS_ENTREGA = Trim(dato)
End Property
'*****************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE MUESTRA      *
'*****************************************************
Public Property Get getID_PROVEEDOR() As Long
    getID_PROVEEDOR = ID_PROVEEDOR
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getDIRECCION() As String
    getDIRECCION = DIRECCION
End Property
Public Property Get getCOD_POSTAL() As Long
    getCOD_POSTAL = COD_POSTAL
End Property
Public Property Get getCIF() As String
    getCIF = CIF
End Property
Public Property Get getTELEFONO() As String
    getTELEFONO = TELEFONO
End Property
Public Property Get getFAX() As String
    getFAX = FAX
End Property
Public Property Get getRESPONSABLE() As String
    getRESPONSABLE = responsable
End Property
Public Property Get getCARGO() As String
    getCARGO = CARGO
End Property
Public Property Get getTIPO() As String
    getTIPO = tipo
End Property
Public Property Get getEMAIL() As String
    getEMAIL = EMAIL
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getCC() As String
    getCC = CC
End Property
Public Property Get getCC_RETENCION() As String
    getCC_RETENCION = CC_RETENCION
End Property
Public Property Get getWEB() As String
    getWEB = WEB
End Property
Public Property Get getMUNICIPIO_ID() As Long
    getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Property Get getINTRA() As Integer
    getINTRA = INTRA
End Property
Public Property Get getEXTRA() As Integer
    getEXTRA = EXTRA
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getPAIS_ID() As Long
    getPAIS_ID = PAIS_ID
End Property
Public Property Get getFP_ID() As Long
    getFP_ID = FP_ID
End Property
Public Property Get getVENCIMIENTO_ID() As Long
    getVENCIMIENTO_ID = VENCIMIENTO_ID
End Property
Public Property Get getBANCO_ID() As Long
    getBANCO_ID = BANCO_ID
End Property
'E0200-I
Public Property Get getES_SUBCONTRATA() As Long
    getES_SUBCONTRATA = ES_SUBCONTRATA
End Property
'E0200-F
'MXXXX-I
Public Property Get getES_FORMADOR() As Long
    getES_FORMADOR = ES_FORMADOR
End Property
'MXXXX-F
'M1334-I
Public Property Get getES_CUENTA_ESPANYA() As Long
    getES_CUENTA_ESPANYA = ES_CUENTA_ESPANYA
End Property
Public Property Get getCUENTA_BANCARIA() As String
    getCUENTA_BANCARIA = CUENTA_BANCARIA
End Property
Public Property Get getBANCO() As String
    getBANCO = banco
End Property
'M1334-F
'M1339-I
Public Property Get getEMAIL_FACTURACION() As String
    getEMAIL_FACTURACION = EMAIL_FACTURACION
End Property
Public Property Get getEVAL_MA() As Integer
    getEVAL_MA = EVAL_MA
End Property
'M1339-F
Public Property Get getDIAS_ENTREGA() As Integer
    getDIAS_ENTREGA = DIAS_ENTREGA
End Property

'*****************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE MUESTRA  *
'*****************************************************
Public Function Carga(ID_PROVEEDOR As Long) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM proveedores WHERE ID_PROVEEDOR=" & ID_PROVEEDOR
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_PROVEEDOR = rs.Fields("ID_PROVEEDOR")
            nombre = rs.Fields("NOMBRE") & ""
            DIRECCION = rs.Fields("DIRECCION") & ""
            If Not IsNull(rs.Fields("COD_POSTAL")) Then
                COD_POSTAL = rs.Fields("COD_POSTAL")
            Else
                COD_POSTAL = 0
            End If
            CIF = rs.Fields("CIF") & ""
            TELEFONO = rs.Fields("TELEFONO") & ""
            FAX = rs.Fields("FAX") & ""
            responsable = rs.Fields("RESPONSABLE") & ""
            CARGO = rs.Fields("CARGO") & ""
            tipo = rs.Fields("TIPO") & ""
            EMAIL = rs.Fields("EMAIL") & ""
            OBSERVACIONES = rs.Fields("OBSERVACIONES") & ""
            CC = rs("CC")
            CC_RETENCION = rs("CC_RETENCION")
            WEB = rs.Fields("WEB") & ""
            anulado = rs("ANULADO")
            INTRA = rs("INTRA")
            EXTRA = rs("EXTRA")
            If IsNull(rs.Fields("MUNICIPIO_ID")) Then
               MUNICIPIO_ID = 0
            Else
               MUNICIPIO_ID = rs.Fields("MUNICIPIO_ID")
            End If
            If IsNull(rs.Fields("PROVINCIA_ID")) Then
               PROVINCIA_ID = 0
            Else
               PROVINCIA_ID = rs.Fields("PROVINCIA_ID")
            End If
            If IsNull(rs.Fields("PAIS_ID")) Then
               PAIS_ID = 0
            Else
               PAIS_ID = rs.Fields("PAIS_ID")
            End If
            FP_ID = rs("FP_ID")
            VENCIMIENTO_ID = rs("VENCIMIENTO_ID")
            BANCO_ID = rs("BANCO_ID")
            'E0200-I
            ES_SUBCONTRATA = rs("ES_SUBCONTRATA")
            'E0200-F
            'MXXXX-I
            ES_FORMADOR = rs("ES_FORMADOR")
            'MXXXX-F
            'M1334-I
            ES_CUENTA_ESPANYA = rs("ES_CUENTA_ESPANYA")
            CUENTA_BANCARIA = rs("CUENTA_BANCARIA")
            banco = rs("BANCO")
            'M1334-F
            'M1339-I
            If Not IsNull(rs("EMAIL_FACTURACION")) Then
                EMAIL_FACTURACION = rs("EMAIL_FACTURACION")
            Else
                EMAIL_FACTURACION = ""
            End If
            If Not IsNull(rs("EVAL_MA")) Then
                EVAL_MA = rs("EVAL_MA")
            Else
                EVAL_MA = -1
            End If
            DIAS_ENTREGA = rs("DIAS_ENTREGA")
            'M1339-F
            Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos del proveedor (Carga)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
    Dim consulta As String
    On Error GoTo fallo:
    If duplicado = False Then
        Me.CrearID_PROVEEDOR
        consulta = "INSERT INTO PROVEEDORES(ID_PROVEEDOR,NOMBRE,DIRECCION," & _
                   "COD_POSTAL,CIF,TELEFONO," & _
                   "FAX,RESPONSABLE,CARGO," & _
                   "TIPO,EMAIL,ANULADO,INTRA,EXTRA," & _
                   "OBSERVACIONES,CC,CC_RETENCION,WEB," & _
                   "MUNICIPIO_ID,PROVINCIA_ID,PAIS_ID,FP_ID,VENCIMIENTO_ID,BANCO_ID,ES_SUBCONTRATA,ES_FORMADOR, ES_CUENTA_ESPANYA, CUENTA_BANCARIA, BANCO, EMAIL_FACTURACION,DIAS_ENTREGA)" & _
                   "values(" & _
                   ID_PROVEEDOR & ",'" & nombre & "','" & DIRECCION & "'," & _
                   COD_POSTAL & ",'" & CIF & "','" & TELEFONO & "','" & _
                   FAX & "','" & responsable & "','" & CARGO & "'," & _
                   IIf(Trim(tipo) = "", "0", "1") & ",'" & EMAIL & "'," & anulado & "," & INTRA & "," & EXTRA & ",'" & _
                   OBSERVACIONES & "','" & CC & "','" & CC_RETENCION & "','" & WEB & "'," & _
                   MUNICIPIO_ID & "," & PROVINCIA_ID & "," & PAIS_ID & "," & FP_ID & "," & VENCIMIENTO_ID & "," & BANCO_ID & "," & ES_SUBCONTRATA & "," & ES_FORMADOR & "," & ES_CUENTA_ESPANYA & ",'" & CUENTA_BANCARIA & "','" & banco & "','" & EMAIL_FACTURACION & "'," & DIAS_ENTREGA & ")"
        execute_bd consulta
    End If
    Insertar = ID_PROVEEDOR
'    MsgBox "Proveedor almacenado correctamente. Código : " & ID_PROVEEDOR, vbInformation, App.Title
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar el proveedor en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearID_PROVEEDOR()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PROVEEDOR) FROM PROVEEDORES"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PROVEEDOR = 1
    Else
        ID_PROVEEDOR = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Modificar() As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE PROVEEDORES SET " & _
               "NOMBRE = '" & nombre & "'" & _
               ",DIRECCION = '" & DIRECCION & "',COD_POSTAL = " & COD_POSTAL & _
               ",CIF = '" & CIF & "',TELEFONO = '" & TELEFONO & "'" & _
               ",FAX = '" & FAX & "',RESPONSABLE = '" & responsable & "'" & _
               ",CARGO = '" & CARGO & "',TIPO = '" & tipo & "'" & _
               ",EMAIL = '" & EMAIL & "'" & _
               ",OBSERVACIONES = '" & OBSERVACIONES & "'" & _
               ",CC = '" & CC & "'" & _
               ",CC_RETENCION = '" & CC_RETENCION & "'" & _
               ",WEB = '" & WEB & "'" & _
               ",MUNICIPIO_ID = " & MUNICIPIO_ID & ",PROVINCIA_ID = " & PROVINCIA_ID & ",PAIS_ID = " & PAIS_ID & _
               ",INTRA = " & INTRA & _
               ",EXTRA = " & EXTRA & _
               ",ANULADO = " & anulado & _
               ",FP_ID = " & FP_ID & _
               ",VENCIMIENTO_ID = " & VENCIMIENTO_ID & _
               ",BANCO_ID = " & BANCO_ID & _
               ",ES_SUBCONTRATA = " & ES_SUBCONTRATA & _
               ",ES_FORMADOR = " & ES_FORMADOR & _
               ",ES_CUENTA_ESPANYA = " & ES_CUENTA_ESPANYA & _
               ",CUENTA_BANCARIA = '" & CUENTA_BANCARIA & "',BANCO = '" & banco & "'" & _
               ",EMAIL_FACTURACION = '" & EMAIL_FACTURACION & "'" & _
               ",DIAS_ENTREGA = " & DIAS_ENTREGA & _
               " WHERE ID_PROVEEDOR = " & ID_PROVEEDOR
               execute_bd consulta
    Modificar = True
    MsgBox "Proveedor modificado correctamente. Código : " & ID_PROVEEDOR, vbInformation, App.Title
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el proveedor en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Eliminar() As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
    On Error GoTo fallo:
    consulta = "SELECT * FROM TIPOS_BOTE_EX WHERE PROVEEDOR_ID = " & ID_PROVEEDOR
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen Tipos de Botes de Reactivos Externos con dicho proveedor. No se puede eliminar.", vbCritical, App.Title
        Exit Function
    End If
    'e0080-i
    ' se comprueba si hay equipos de dicho proveedor
    Set rs = Nothing
    consulta = "SELECT * FROM EQUIPOS WHERE PROVEEDOR_ID = " & ID_PROVEEDOR
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen equipos del proveedor seleccionado. No se puede eliminar.", vbCritical, App.Title
        Exit Function
    End If
    ' se comprueba si hay calibraciones de equipos de dicho proveedor
    Set rs = Nothing
    consulta = "SELECT * FROM EQUIPOS WHERE CALIBRADOR_ID = " & ID_PROVEEDOR
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen calibraciones de equipos del proveedor seleccionado. No se puede eliminar.", vbCritical, App.Title
        Exit Function
    End If
    Set rs = Nothing
    consulta = "SELECT * FROM SC_PAQUETES WHERE SUBCONTRATA_ID = " & ID_PROVEEDOR
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen SUBCONTRATACIONES de ensayos del proveedor seleccionado. No se puede eliminar.", vbCritical, App.Title
        Exit Function
    End If
    Set rs = Nothing
    'e0080-f
    consulta = "DELETE FROM PROVEEDORES " & _
               " WHERE " & _
               " ID_PROVEEDOR = " & ID_PROVEEDOR
               execute_bd consulta
    Eliminar = True
    MsgBox "Proveedor eliminado correctamente. Código : " & ID_PROVEEDOR, vbInformation, App.Title
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el proveedor de la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Function Anular(ID As Long) As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
    On Error GoTo fallo:
    consulta = "SELECT ID FROM PROVEEDORES_FACTURAS WHERE PROVEEDOR_ID = " & ID & " AND F_PAGO = '0000-00-00'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "El proveedor no se puede anular, tiene facturas pendientes de Pago.", vbCritical, App.Title
        Exit Function
    End If
        
    consulta = "UPDATE PROVEEDORES SET ANULADO = 1 " & _
               " WHERE " & _
               " ID_PROVEEDOR = " & ID
    execute_bd consulta
    Anular = True
    MsgBox "Proveedor ANULADO correctamente. Código : " & ID, vbInformation, App.Title
    Exit Function
fallo:
    Anular = False
    MsgBox "Error al ANULAR el proveedor de la bd : " & Err.Description, vbCritical, Err.Number
End Function

'E0200-I
'MXXXX-I
'Public Function Listado(strMostrarSoloSubcontratas As String, lNombre As String, lCif As String, lTelefono As String, servicio As String, booSinServicio As Boolean, booAsigEquipos As Boolean, booAsigReactivos As Boolean, booNoAsig As Boolean, booanulados As Boolean) As ADODB.Recordset
Public Function Listado(strMostrarSoloSubcontratas As String, lNombre As String, lCif As String, lTelefono As String, servicio As String, booSinServicio As Boolean, booAsigEquipos As Boolean, booAsigReactivos As Boolean, booNoAsig As Boolean, booanulados As Boolean, booformadores As Boolean, booIntra As Boolean, booExtra As Boolean) As ADODB.Recordset
'MXXXX-F
    Dim consulta As String
    Dim filtro As String
    If booIntra = True Then
        filtro = filtro & " AND P.INTRA = 1 "
    End If
    If booExtra = True Then
        filtro = filtro & " AND P.EXTRA = 1 "
    End If
'    booAsigEquipos As Boolean, booAsigReactivos As Boolean, booNoAsig
    If booSinServicio = True Then ' Sin servicio asignado
            consulta = "SELECT * " & _
                       "  FROM PROVEEDORES P LEFT JOIN PROVEEDORES_SERVICIOS PS ON P.ID_PROVEEDOR = PS.PROVEEDOR_ID " & _
                       " WHERE P.ID_PROVEEDOR <> 0 " & _
                       IIf(strMostrarSoloSubcontratas = "1", " AND P.ES_SUBCONTRATA = 1 ", "") & _
                       "   AND P.NOMBRE LIKE '%" & Trim(lNombre) & "%' " & _
                       "   AND P.CIF LIKE '%" & Trim(lCif) & "%' " & _
                       "   AND P.TELEFONO LIKE '%" & Trim(lTelefono) & "%' " & _
                       filtro & _
                       " AND PS.SERVICIO_ID IS NULL " & _
                       IIf(booanulados = False, " AND P.ANULADO=0 ", "") & _
                       " ORDER BY P.NOMBRE"
    
    ElseIf booAsigEquipos = True Then
        consulta = "SELECT DISTINCT P.* " & _
                   "  FROM PROVEEDORES P  " & _
                   " INNER JOIN EQUIPOS EQ ON P.ID_PROVEEDOR = EQ.PROVEEDOR_ID  " & _
                   " where ES_SUBCONTRATA = 0 And p.ID_PROVEEDOR <> 0  " & _
                       "   AND P.NOMBRE LIKE '%" & Trim(lNombre) & "%' " & _
                       "   AND P.CIF LIKE '%" & Trim(lCif) & "%' " & _
                       "   AND P.TELEFONO LIKE '%" & Trim(lTelefono) & "%' " & _
                       filtro & _
                       IIf(booanulados = False, " AND P.ANULADO=0 ", "") & _
                   " ORDER BY P.NOMBRE"
    ElseIf booAsigReactivos = True Then
        consulta = "SELECT DISTINCT P.* " & _
                   "  FROM PROVEEDORES P  " & _
                   " INNER JOIN TIPOS_BOTE_EX EQ ON P.ID_PROVEEDOR = EQ.PROVEEDOR_ID  " & _
                   " where ES_SUBCONTRATA = 0 And p.ID_PROVEEDOR <> 0  " & _
                       "   AND P.NOMBRE LIKE '%" & Trim(lNombre) & "%' " & _
                       "   AND P.CIF LIKE '%" & Trim(lCif) & "%' " & _
                       "   AND P.TELEFONO LIKE '%" & Trim(lTelefono) & "%' " & _
                       filtro & _
                       IIf(booanulados = False, " AND P.ANULADO=0 ", "") & _
                   " ORDER BY P.NOMBRE"
    ElseIf booNoAsig = True Then
        consulta = " SELECT DISTINCT P.* " & _
                   "   FROM PROVEEDORES P  " & _
                   "   LEFT JOIN EQUIPOS EQ ON P.ID_PROVEEDOR = EQ.PROVEEDOR_ID  " & _
                   "   LEFT JOIN TIPOS_BOTE_EX TR ON P.ID_PROVEEDOR = TR.PROVEEDOR_ID  " & _
                   "  where eq.PROVEEDOR_ID Is Null " & _
                   "    AND TR.PROVEEDOR_ID IS NULL " & _
                   "    AND ES_SUBCONTRATA = 0 " & _
                       "   AND P.NOMBRE LIKE '%" & Trim(lNombre) & "%' " & _
                       "   AND P.CIF LIKE '%" & Trim(lCif) & "%' " & _
                       "   AND P.TELEFONO LIKE '%" & Trim(lTelefono) & "%' " & _
                       filtro & _
                       IIf(booanulados = False, " AND P.ANULADO=0 ", "") & _
                   "  ORDER BY P.NOMBRE "
    Else
        If servicio = "" Or servicio = "0" Then
            consulta = "SELECT * " & _
                       "  FROM PROVEEDORES P " & _
                       " WHERE ID_PROVEEDOR <> 0 " & _
                       IIf(strMostrarSoloSubcontratas = "1", " AND ES_SUBCONTRATA = 1 ", "") & _
                       "   AND NOMBRE LIKE '%" & lNombre & "%' " & _
                       "   AND CIF LIKE '%" & lCif & "%' " & _
                       "   AND TELEFONO LIKE '%" & lTelefono & "%' " & _
                       filtro & _
                       IIf(booanulados = False, " AND ANULADO=0 ", "") & _
                       " ORDER BY NOMBRE"
        Else
            consulta = "SELECT * " & _
                       "  FROM PROVEEDORES P INNER JOIN PROVEEDORES_SERVICIOS PS ON P.ID_PROVEEDOR = PS.PROVEEDOR_ID " & _
                       " WHERE P.ID_PROVEEDOR <> 0 " & _
                       IIf(strMostrarSoloSubcontratas = "1", " AND P.ES_SUBCONTRATA = 1 ", "") & _
                       "   AND P.NOMBRE LIKE '%" & Trim(lNombre) & "%' " & _
                       "   AND P.CIF LIKE '%" & Trim(lCif) & "%' " & _
                       "   AND P.TELEFONO LIKE '%" & Trim(lTelefono) & "%' " & _
                       " AND PS.SERVICIO_ID = " & CInt(servicio) & _
                       filtro & _
                       IIf(booanulados = False, " AND P.ANULADO=0 ", "") & _
                       " ORDER BY P.NOMBRE"
        
        End If
    End If
    Set Listado = datos_bd(consulta)
End Function
'E0200-F

Public Function duplicado() As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PROVEEDOR FROM PROVEEDORES WHERE NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        duplicado = False
    Else
        ID_PROVEEDOR = rs(0)
        duplicado = True
    End If
    Exit Function
fallo:
    duplicado = False
End Function

Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PROVEEDOR,NOMBRE FROM PROVEEDORES order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_con_reactivos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT P.ID_PROVEEDOR AS A1,P.NOMBRE AS A2" & _
               "  FROM PROVEEDORES P, TIPOS_BOTE_EX T " & _
               " WHERE P.ID_PROVEEDOR = T.PROVEEDOR_ID " & _
               " ORDER BY NOMBRE"
    Set Listado_con_reactivos = datos_bd(consulta)
End Function

'E0046-I
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setFILTRO = filtro
        .setTABLA = "PROVEEDORES"
        .setDESCRIPCION = "Proveedores"
        .setPK = "ID_PROVEEDOR"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
'E0046-F

'E0200-I
Public Function listado_subcontratas() As ADODB.Recordset
    Dim consulta As String

    consulta = "SELECT ID_PROVEEDOR, NOMBRE " & _
               "  FROM PROVEEDORES " & _
               " WHERE ES_SUBCONTRATA = 1 " & _
               " ORDER BY NOMBRE "
    Set listado_subcontratas = datos_bd(consulta)
End Function
'E0200-F
Public Function Evaluacion(ID_PROVEEDOR As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE PROVEEDORES SET " & _
               " EVAL_MA = " & EVAL_MA & _
               " WHERE ID_PROVEEDOR = " & ID_PROVEEDOR
    execute_bd consulta
    Evaluacion = True
    MsgBox "Evaluación informada correctamente.", vbInformation, App.Title
    Exit Function
fallo:
    Evaluacion = False
    MsgBox "Error al modificar el proveedor en la bd : " & Err.Description, vbCritical, Err.Number
End Function


