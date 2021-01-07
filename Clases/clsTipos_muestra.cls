VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_muestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE tipos_muestra      *
'***********************************************************
Private id_tipo_muestra As Integer
Private CODIGO As String
Private nombre As String
Private SECTOR_ID As Integer
Private FAMILIA_ID As Integer
Private TIPO_DOCUMENTO_ID As Integer
Private TIPO_ESPECIAL_ID As Integer
Private TIPO_ENSAYO_ID As Integer
Private RESPONSABLE_ID As Integer
Private REQUIERE_PRODUCTO As Integer
Private ENAC As Integer
Private nadcap As Integer
Private anulado As Integer
Public Function ImprimirListado() As Boolean
    Dim objfrm As New frmReport
    With objfrm
        .iniciar
        .informe = "muestras\rptListadoTiposMuestra"
        
        .criterio = ""
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objfrm
    Set objfrm = Nothing
End Function

Public Property Let setID_TIPO_MUESTRA(ByVal dato As Integer)
    id_tipo_muestra = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
    CODIGO = Trim(dato)
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setSECTOR_ID(ByVal dato As Integer)
    SECTOR_ID = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Integer)
    FAMILIA_ID = dato
End Property
Public Property Let setTIPO_DOCUMENTO_ID(ByVal dato As Integer)
    TIPO_DOCUMENTO_ID = dato
End Property
Public Property Let setTIPO_ESPECIAL_ID(ByVal dato As Integer)
    TIPO_ESPECIAL_ID = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Integer)
    TIPO_ENSAYO_ID = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Integer)
    RESPONSABLE_ID = dato
End Property
Public Property Let setREQUIERE_PRODUCTO(ByVal dato As Integer)
    REQUIERE_PRODUCTO = dato
End Property
Public Property Let setENAC(ByVal dato As Integer)
    ENAC = dato
End Property
Public Property Let setNADCAP(ByVal dato As Integer)
    nadcap = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
Public Property Get getID_TIPO_MUESTRA() As Integer
    getID_TIPO_MUESTRA = id_tipo_muestra
End Property
Public Property Get getCODIGO() As String
    getCODIGO = CODIGO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getSECTOR_ID() As Integer
    getSECTOR_ID = SECTOR_ID
End Property
Public Property Get getFAMILIA_ID() As Integer
    getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getTIPO_DOCUMENTO_ID() As Integer
    getTIPO_DOCUMENTO_ID = TIPO_DOCUMENTO_ID
End Property
Public Property Get getTIPO_ESPECIAL_ID() As Integer
    getTIPO_ESPECIAL_ID = TIPO_ESPECIAL_ID
End Property
Public Property Get getTIPO_ENSAYO_ID() As Integer
    getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getRESPONSABLE_ID() As Integer
    getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getREQUIERE_PRODUCTO() As Integer
    getREQUIERE_PRODUCTO = REQUIERE_PRODUCTO
End Property
Public Property Get getENAC() As Integer
    getENAC = ENAC
End Property
Public Property Get getNADCAP() As Integer
    getNADCAP = nadcap
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_muestra where anulado = 0 order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_sin_anular() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_tipo_muestra,concat(nombre,' (',codigo,')') as nombre FROM tipos_muestra where anulado = 0 order by nombre"
    Set Listado_sin_anular = datos_bd(consulta)
End Function
Public Function Listado_todas() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_tipo_muestra,concat(nombre,' (',codigo,')') as nombre FROM tipos_muestra order by nombre"
    Set Listado_todas = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_MUESTRA,NOMBRE FROM tipos_muestra where anulado = 0 order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_por_familias(familia As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_muestra where familia_id=" & familia & " order by nombre"
    Set Listado_por_familias = datos_bd(consulta)
End Function

Public Function Eliminar(ID_TIPO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "update tipos_muestra set anulado=1 where id_tipo_muestra=" & ID_TIPO
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el tipo de muestra", vbCritical, Err.Description
End Function

Public Function Insertar() As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into tipos_muestra " & _
               " values(" & _
               id_tipo_muestra & ",'" & UCase(CODIGO) & "','" & nombre & "'," & _
               SECTOR_ID & "," & FAMILIA_ID & "," & TIPO_DOCUMENTO_ID & "," & _
               TIPO_ESPECIAL_ID & "," & TIPO_ENSAYO_ID & ",0," & RESPONSABLE_ID & "," & REQUIERE_PRODUCTO & "," & ENAC & "," & nadcap & ")"
    execute_bd consulta
    Insertar = True
    Exit Function
fallo:
    Insertar = False
    MsgBox "Error al insertar el tipo de muestra.", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_MUESTRA) FROM tipos_muestra"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        id_tipo_muestra = 1
    Else
        id_tipo_muestra = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Function Modificar(ID_TIPO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "update tipos_muestra " & _
               " set CODIGO = '" & UCase(CODIGO) & "'," & _
               " NOMBRE = '" & nombre & "'," & _
               " SECTOR_ID = " & SECTOR_ID & "," & _
               " FAMILIA_ID = " & FAMILIA_ID & "," & _
               " TIPO_DOCUMENTO_ID = " & TIPO_DOCUMENTO_ID & "," & _
               " RESPONSABLE_ID = " & RESPONSABLE_ID & "," & _
               " REQUIERE_PRODUCTO = " & REQUIERE_PRODUCTO & "," & _
               " ENAC = " & ENAC & "," & _
               " NADCAP = " & nadcap & "," & _
               " TIPO_ESPECIAL_ID = " & TIPO_ESPECIAL_ID & "," & _
               " TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID & _
               " WHERE ID_TIPO_MUESTRA = " & ID_TIPO
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al modificar el tipo de muestra.", vbCritical, Err.Description
End Function
Public Function cargar(TIPO_MUESTRA As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_muestra WHERE id_tipo_muestra=" & TIPO_MUESTRA
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        id_tipo_muestra = rs("ID_TIPO_MUESTRA")
        nombre = rs("NOMBRE")
        CODIGO = rs("CODIGO")
        SECTOR_ID = rs("SECTOR_ID")
        FAMILIA_ID = rs("FAMILIA_ID")
        TIPO_DOCUMENTO_ID = rs("TIPO_DOCUMENTO_ID")
        TIPO_ESPECIAL_ID = rs("TIPO_ESPECIAL_ID")
        TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
        RESPONSABLE_ID = rs("RESPONSABLE_ID")
        REQUIERE_PRODUCTO = rs("REQUIERE_PRODUCTO")
        ENAC = rs("ENAC")
        nadcap = rs("NADCAP")
        anulado = rs("ANULADO")
        cargar = True
    Else
        cargar = False
    End If
    Exit Function
fallo:
    cargar = False
End Function
Public Function Listado_no_especiales() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_muestra where tipo_especial_id = 0 order by nombre"
    Set Listado_no_especiales = datos_bd(consulta)
End Function

Public Function Listado_especiales() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_muestra where tipo_especial_id <> 0 order by nombre"
    Set Listado_especiales = datos_bd(consulta)
End Function

Public Function lista(id_tipo_muestra As Long, Optional filtro As String, Optional tipo_especial As Integer, Optional producto As Boolean, Optional ENAC As Boolean, Optional nadcap As Boolean, Optional tipoENAC As Integer, Optional tipoEnsayo As Integer) As ADODB.Recordset
    Dim consulta As String
    Dim f As String
    If Trim(filtro) <> "" Then
       f = f & " AND tm.nombre like '%" & filtro & "%' "
    End If
    If tipo_especial <> 0 Then
        f = f & " AND TM.TIPO_ESPECIAL_ID = " & tipo_especial
    End If
    If producto = True Then
        f = f & " AND TM.REQUIERE_PRODUCTO = 1 "
    End If
    If ENAC = True Then
        f = f & " AND TM.ENAC = " & tipoENAC
    End If
    If nadcap = True Then
        f = f & " AND TM.NADCAP = 1 "
    End If
    If tipoEnsayo <> 0 Then
        f = f & " AND TM.TIPO_ENSAYO_ID = " & tipoEnsayo
    End If
    If id_tipo_muestra <> 0 Then
        f = " AND tm.id_tipo_muestra = " & id_tipo_muestra
    End If
    consulta = "SELECT tm.codigo,tm.nombre,te.nombre,tensayo.descripcion,s.nombre,f.nombre,tm.id_tipo_muestra,tm.anulado " & _
               "FROM tipos_muestra tm, sectores s, familias f, tipos_especial te, decodificadora tensayo " & _
               "WHERE tm.sector_id = s.id_sector " & _
               "  AND tm.familia_id = f.id_familia " & _
               "  AND tm.tipo_Especial_id = te.id_tipo_especial " & _
               "  AND tensayo.valor = tm.tipo_ensayo_id and tensayo.codigo = " & DECODIFICADORA.DECODIFICADORA_TM_TIPOS_ENSAYOS & _
               f & _
               " order by tm.nombre"
    Set lista = datos_bd(consulta)
End Function
Public Function Comprobar_codigo(CODIGO) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_MUESTRA FROM tipos_muestra where codigo = '" & UCase(CODIGO) & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Comprobar_codigo = False
    Else
        Comprobar_codigo = True
    End If
    Set rs = Nothing
End Function
Public Function Listado_Fluidos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_muestra where tipo_especial_id = 4 order by nombre"
    Set Listado_Fluidos = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TIPOS_MUESTRA"
        .setDESCRIPCION = "Tipos Muestra"
        .setPK = "ID_TIPO_MUESTRA"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        .setFILTRO = filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function
