VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_EMPLEADO As Long
Private nombre As String
Private APELLIDOS As String
Private USUARIO As String
Private PassWord As String
Private PER_IMPRESION As Integer
Private PER_FACTURACION As Integer
Private PER_MODIFICACION As Integer
Private PER_ELIMINACION As Integer
Private PER_USUARIOS As Integer
Private PER_EDICION As Integer
Private PER_CIERRE As Integer
Private PER_PEDIDOS_REACTIVOS As Integer
Private PER_EMPLEADOS As Integer
Private PER_CONTABILIDAD As Integer
Private PER_MATRIZ_CUALIF As Integer
Private PER_ENVIO_PEDIDOS_PROVEEDOR As Integer
Private PER_DOCUMENTACION_CALIDAD As Integer
Private PER_PROYECTOS As Integer
Private PER_NC As Integer
Private PER_DATOS_ESPECIALES As Integer
Private PER_PNT As Integer
Private PER_ADMIN_PNT As Integer
Private uso As String
Private firma As String
Private imagen As String
Private anulado As Integer
Private EMAIL As String
Private FNMT_RUTA As String
Private FNMT_PASS As String
Private RESPONSABLE_DEPARTAMENTOS(12) As Integer
Private RESPONSABLE_DEPARTAMENTOS_INICIAL(12) As Integer
Private DEPARTAMENTOS_USUARIO As New clsGenericCollection
Private PER_VIDEOS As Integer
Private PER_IMPRESION_PNT As Integer
Private PER_INCURRIDOS As Integer
Private PER_REVISION As Integer
Private PER_OFERTAS As Integer
Private PER_PLAZO_ENTREGA_CAMBIO As Integer
Private PER_PLAZO_ENTREGA_LISTADO As Integer
Private PER_INDICADORES_CLIENTE As Integer
Private PER_FAMILIAS_CA As Integer
Private REVISION_CARGO As String
'M1144-I
Private PER_TRAMITACION_CONTRATA As Integer
Private PER_RFI As Integer
Private PER_PFA As Integer
'M1144-F
Private PER_SCG As Integer
Private PER_TESORERIA_MENU As Integer
Private PER_TESORERIA_FP As Integer
'M1377-I
Private PER_NORMAS_NO_CONTROLADAS As Integer
Private PER_DES_PRODUCTO As Integer
Private PER_MOD_CLIENTE As Integer
Private PER_INCIDENCIAS As Integer
Private PER_PRODUCTIVIDAD As Integer
'M1377-F
Public Property Let setID_EMPLEADO(ByVal dato As Long)
    ID_EMPLEADO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setAPELLIDOS(ByVal dato As String)
    APELLIDOS = Trim(dato)
End Property
Public Property Let setUSUARIO(ByVal dato As String)
    USUARIO = Trim(dato)
End Property
Public Property Let setPASSWORD(ByVal dato As String)
    PassWord = Trim(dato)
End Property
Public Property Let setPER_IMPRESION(ByVal dato As Boolean)
    If dato = True Then
        PER_IMPRESION = 1
    Else
        PER_IMPRESION = 0
    End If
End Property
Public Property Let setPER_FACTURACION(ByVal dato As Boolean)
    If dato = True Then
        PER_FACTURACION = 1
    Else
        PER_FACTURACION = 0
    End If
End Property
Public Property Let setPER_MODIFICACION(ByVal dato As Boolean)
    If dato = True Then
        PER_MODIFICACION = 1
    Else
        PER_MODIFICACION = 0
    End If
End Property
Public Property Let setPER_ELIMINACION(ByVal dato As Boolean)
    If dato = True Then
        PER_ELIMINACION = 1
    Else
        PER_ELIMINACION = 0
    End If
End Property
Public Property Let setPER_USUARIOS(ByVal dato As Boolean)
    If dato = True Then
        PER_USUARIOS = 1
    Else
        PER_USUARIOS = 0
    End If
End Property
Public Property Let setPER_EDICION(ByVal dato As Boolean)
    If dato = True Then
        PER_EDICION = 1
    Else
        PER_EDICION = 0
    End If
End Property
Public Property Let setPER_CIERRE(ByVal dato As Boolean)
    If dato = True Then
        PER_CIERRE = 1
    Else
        PER_CIERRE = 0
    End If
End Property
Public Property Let setPER_PEDIDOS_REACTIVOS(ByVal dato As Boolean)
    If dato = True Then
        PER_PEDIDOS_REACTIVOS = 1
    Else
        PER_PEDIDOS_REACTIVOS = 0
    End If
End Property
Public Property Let setPER_EMPLEADOS(ByVal dato As Boolean)
    If dato = True Then
        PER_EMPLEADOS = 1
    Else
        PER_EMPLEADOS = 0
    End If
End Property
Public Property Let setPER_CONTABILIDAD(ByVal dato As Boolean)
    If dato = True Then
        PER_CONTABILIDAD = 1
    Else
        PER_CONTABILIDAD = 0
    End If
End Property
Public Property Let setPER_MATRIZ_CUALIF(ByVal dato As Boolean)
    If dato = True Then
        PER_MATRIZ_CUALIF = 1
    Else
        PER_MATRIZ_CUALIF = 0
    End If
End Property
Public Property Let setPER_ENVIO_PEDIDOS_PROVEEDOR(ByVal dato As Boolean)
    If dato = True Then
        PER_ENVIO_PEDIDOS_PROVEEDOR = 1
    Else
        PER_ENVIO_PEDIDOS_PROVEEDOR = 0
    End If
End Property
Public Property Let setPER_DOCUMENTACION_CALIDAD(ByVal dato As Boolean)
    If dato = True Then
        PER_DOCUMENTACION_CALIDAD = 1
    Else
        PER_DOCUMENTACION_CALIDAD = 0
    End If
End Property
Public Property Let setPER_PROYECTOS(ByVal dato As Boolean)
    If dato = True Then
        PER_PROYECTOS = 1
    Else
        PER_PROYECTOS = 0
    End If
End Property
Public Property Let setPER_NC(ByVal dato As Boolean)
    If dato = True Then
        PER_NC = 1
    Else
        PER_NC = 0
    End If
End Property
Public Property Let setPER_DATOS_ESPECIALES(ByVal dato As Boolean)
    If dato = True Then
        PER_DATOS_ESPECIALES = 1
    Else
        PER_DATOS_ESPECIALES = 0
    End If
End Property

Public Property Let setPER_VIDEOS(ByVal dato As Boolean)
    If dato = True Then
        PER_VIDEOS = 1
    Else
        PER_VIDEOS = 0
    End If
End Property
Public Property Let setPER_IMPRESION_PNT(ByVal dato As Boolean)
    If dato = True Then
        PER_IMPRESION_PNT = 1
    Else
        PER_IMPRESION_PNT = 0
    End If
End Property
Public Property Let setPER_INCURRIDOS(ByVal dato As Boolean)
    If dato = True Then
        PER_INCURRIDOS = 1
    Else
        PER_INCURRIDOS = 0
    End If
End Property
Public Property Let setREVISION_CARGO(ByVal dato As String)
    REVISION_CARGO = dato
End Property
Public Property Let setPER_REVISION(ByVal dato As Boolean)
    If dato = True Then
        PER_REVISION = 1
    Else
        PER_REVISION = 0
    End If
End Property
Public Property Let setPER_OFERTAS(ByVal dato As Boolean)
    If dato = True Then
        PER_OFERTAS = 1
    Else
        PER_OFERTAS = 0
    End If
End Property
Public Property Let setPER_PLAZO_ENTREGA_CAMBIO(ByVal dato As Boolean)
    If dato = True Then
        PER_PLAZO_ENTREGA_CAMBIO = 1
    Else
        PER_PLAZO_ENTREGA_CAMBIO = 0
    End If
End Property
Public Property Let setPER_PLAZO_ENTREGA_LISTADO(ByVal dato As Boolean)
    If dato = True Then
        PER_PLAZO_ENTREGA_LISTADO = 1
    Else
        PER_PLAZO_ENTREGA_LISTADO = 0
    End If
End Property
Public Property Let setPER_INDICADORES_CLIENTE(ByVal dato As Boolean)
    If dato = True Then
        PER_INDICADORES_CLIENTE = 1
    Else
        PER_INDICADORES_CLIENTE = 0
    End If
End Property
Public Property Let setPER_FAMILIAS_CA(ByVal dato As Boolean)
    If dato = True Then
        PER_FAMILIAS_CA = 1
    Else
        PER_FAMILIAS_CA = 0
    End If
End Property

Public Property Let setPER_PNT(ByVal dato As Boolean)
    If dato = True Then
        PER_PNT = 1
    Else
        PER_PNT = 0
    End If
End Property
Public Property Let setPER_ADMIN_PNT(ByVal dato As Boolean)
    If dato = True Then
        PER_ADMIN_PNT = 1
    Else
        PER_ADMIN_PNT = 0
    End If
End Property
Public Property Let setUSO(ByVal dato As String)
    uso = dato
End Property
Public Property Let setFIRMA(ByVal dato As String)
    firma = dato
End Property
Public Property Let setIMAGEN(ByVal dato As String)
    imagen = dato
End Property
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
'E0139-I
Public Property Let setEMAIL(ByVal dato As String)
    EMAIL = Trim(dato)
End Property
Public Property Let setFNMT_RUTA(ByVal dato As String)
    FNMT_RUTA = Trim(dato)
End Property
Public Property Let setFNMT_PASS(ByVal dato As String)
    FNMT_PASS = Trim(dato)
End Property
'E0139-F
'M1144-I
Public Property Let setPER_TRAMITACION_CONTRATA(ByVal dato As Boolean)
    If dato = True Then
        PER_TRAMITACION_CONTRATA = 1
    Else
        PER_TRAMITACION_CONTRATA = 0
    End If
End Property

Public Property Let setPER_RFI(ByVal dato As Boolean)
    If dato = True Then
        PER_RFI = 1
    Else
        PER_RFI = 0
    End If
End Property

Public Property Let setPER_PFA(ByVal dato As Boolean)
    If dato = True Then
        PER_PFA = 1
    Else
        PER_PFA = 0
    End If
End Property
'M1144-F
Public Property Let setPER_SCG(ByVal dato As Boolean)
    If dato = True Then
        PER_SCG = 1
    Else
        PER_SCG = 0
    End If
End Property
Public Property Let setPER_TESORERIA_MENU(ByVal dato As Boolean)
    If dato = True Then
        PER_TESORERIA_MENU = 1
    Else
        PER_TESORERIA_MENU = 0
    End If
End Property
Public Property Let setPER_TESORERIA_FP(ByVal dato As Boolean)
    If dato = True Then
        PER_TESORERIA_FP = 1
    Else
        PER_TESORERIA_FP = 0
    End If
End Property
'M1377-I
Public Property Let setPER_NORMAS_NO_CONTROLADAS(ByVal dato As Boolean)
    If dato = True Then
        PER_NORMAS_NO_CONTROLADAS = 1
    Else
        PER_NORMAS_NO_CONTROLADAS = 0
    End If
End Property
Public Property Let setPER_DES_PRODUCTO(ByVal dato As Boolean)
    If dato = True Then
        PER_DES_PRODUCTO = 1
    Else
        PER_DES_PRODUCTO = 0
    End If
End Property
Public Property Let setPER_MOD_CLIENTE(ByVal dato As Boolean)
    If dato = True Then
        PER_MOD_CLIENTE = 1
    Else
        PER_MOD_CLIENTE = 0
    End If
End Property
Public Property Let setPER_INCIDENCIAS(ByVal dato As Boolean)
    If dato = True Then
        PER_INCIDENCIAS = 1
    Else
        PER_INCIDENCIAS = 0
    End If
End Property
Public Property Let setPER_PRODUCTIVIDAD(ByVal dato As Boolean)
    If dato = True Then
        PER_PRODUCTIVIDAD = 1
    Else
        PER_PRODUCTIVIDAD = 0
    End If
End Property
'M1377-F
Public Property Get getID_EMPLEADO() As Long
    getID_EMPLEADO = ID_EMPLEADO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getAPELLIDOS() As String
    getAPELLIDOS = APELLIDOS
End Property
Public Property Get getUSUARIO() As String
    getUSUARIO = USUARIO
End Property
Public Property Get getPASSWORD() As String
    getPASSWORD = PassWord
End Property
Public Property Get getPER_IMPRESION() As Boolean
    If PER_IMPRESION = 1 Then
        getPER_IMPRESION = True
    Else
        getPER_IMPRESION = False
    End If
End Property
Public Property Get getPER_FACTURACION() As Boolean
    If PER_FACTURACION = 1 Then
        getPER_FACTURACION = True
    Else
        getPER_FACTURACION = False
    End If
End Property
Public Property Get getPER_MODIFICACION() As Boolean
    If PER_MODIFICACION = 1 Then
        getPER_MODIFICACION = True
    Else
        getPER_MODIFICACION = False
    End If
End Property
Public Property Get getPER_ELIMINACION() As Boolean
    If PER_ELIMINACION = 1 Then
        getPER_ELIMINACION = True
    Else
        getPER_ELIMINACION = False
    End If
End Property
Public Property Get getPER_VIDEOS() As Boolean
    If PER_VIDEOS = 1 Then
        getPER_VIDEOS = True
    Else
        getPER_VIDEOS = False
    End If
End Property
Public Property Get getPER_IMPRESION_PNT() As Boolean
    If PER_IMPRESION_PNT = 1 Then
        getPER_IMPRESION_PNT = True
    Else
        getPER_IMPRESION_PNT = False
    End If
End Property
Public Property Get getPER_INCURRIDOS() As Boolean
    If PER_INCURRIDOS = 1 Then
        getPER_INCURRIDOS = True
    Else
        getPER_INCURRIDOS = False
    End If
End Property
Public Property Get getPER_REVISION() As Boolean
    If PER_REVISION = 1 Then
        getPER_REVISION = True
    Else
        getPER_REVISION = False
    End If
End Property
Public Property Get getPER_OFERTAS() As Boolean
    If PER_OFERTAS = 1 Then
        getPER_OFERTAS = True
    Else
        getPER_OFERTAS = False
    End If
End Property
Public Property Get getPER_PLAZO_ENTREGA_CAMBIO() As Boolean
    If PER_PLAZO_ENTREGA_CAMBIO = 1 Then
        getPER_PLAZO_ENTREGA_CAMBIO = True
    Else
        getPER_PLAZO_ENTREGA_CAMBIO = False
    End If
End Property
Public Property Get getPER_PLAZO_ENTREGA_LISTADO() As Boolean
    If PER_PLAZO_ENTREGA_LISTADO = 1 Then
        getPER_PLAZO_ENTREGA_LISTADO = True
    Else
        getPER_PLAZO_ENTREGA_LISTADO = False
    End If
End Property
Public Property Get getPER_INDICADORES_CLIENTE() As Boolean
    If PER_INDICADORES_CLIENTE = 1 Then
        getPER_INDICADORES_CLIENTE = True
    Else
        getPER_INDICADORES_CLIENTE = False
    End If
End Property
Public Property Get getPER_FAMILIAS_CA() As Boolean
    If PER_FAMILIAS_CA = 1 Then
        getPER_FAMILIAS_CA = True
    Else
        getPER_FAMILIAS_CA = False
    End If
End Property

Public Property Get getPER_USUARIOS() As Boolean
    If PER_USUARIOS = 1 Then
        getPER_USUARIOS = True
    Else
        getPER_USUARIOS = False
    End If
End Property
Public Property Get getPER_EDICION() As Boolean
    If PER_EDICION = 1 Then
        getPER_EDICION = True
    Else
        getPER_EDICION = False
    End If
End Property
Public Property Get getPER_CIERRE() As Boolean
    If PER_CIERRE = 1 Then
        getPER_CIERRE = True
    Else
        getPER_CIERRE = False
    End If
End Property
Public Property Get getPER_CONTABILIDAD() As Boolean
    If PER_CONTABILIDAD = 1 Then
        getPER_CONTABILIDAD = True
    Else
        getPER_CONTABILIDAD = False
    End If
End Property

Public Property Get getPER_MATRIZ_CUALIF() As Boolean
    If PER_MATRIZ_CUALIF = 1 Then
        getPER_MATRIZ_CUALIF = True
    Else
        getPER_MATRIZ_CUALIF = False
    End If
End Property
Public Property Get getPER_ENVIO_PEDIDOS_PROVEEDOR() As Boolean
    If PER_ENVIO_PEDIDOS_PROVEEDOR = 1 Then
        getPER_ENVIO_PEDIDOS_PROVEEDOR = True
    Else
        getPER_ENVIO_PEDIDOS_PROVEEDOR = False
    End If
End Property
Public Property Get getPER_DOCUMENTACION_CALIDAD() As Boolean
    If PER_DOCUMENTACION_CALIDAD = 1 Then
        getPER_DOCUMENTACION_CALIDAD = True
    Else
        getPER_DOCUMENTACION_CALIDAD = False
    End If
End Property
Public Property Get getPER_PROYECTOS() As Boolean
    If PER_PROYECTOS = 1 Then
        getPER_PROYECTOS = True
    Else
        getPER_PROYECTOS = False
    End If
End Property
Public Property Get getPER_NC() As Boolean
    If PER_NC = 1 Then
        getPER_NC = True
    Else
        getPER_NC = False
    End If
End Property
Public Property Get getPER_DATOS_ESPECIALES() As Boolean
    If PER_DATOS_ESPECIALES = 1 Then
        getPER_DATOS_ESPECIALES = True
    Else
        getPER_DATOS_ESPECIALES = False
    End If
End Property

Public Property Get getPER_PNT() As Boolean
    If PER_PNT = 1 Then
        getPER_PNT = True
    Else
        getPER_PNT = False
    End If
End Property
Public Property Get getPER_ADMIN_PNT() As Boolean
    If PER_ADMIN_PNT = 1 Then
        getPER_ADMIN_PNT = True
    Else
        getPER_ADMIN_PNT = False
    End If
End Property
Public Property Get getPER_PEDIDOS_REACTIVOS() As Boolean
    If PER_PEDIDOS_REACTIVOS = 1 Then
        getPER_PEDIDOS_REACTIVOS = True
    Else
        getPER_PEDIDOS_REACTIVOS = False
    End If
End Property
Public Property Get getPER_EMPLEADOS() As Boolean
    If PER_EMPLEADOS = 1 Then
        getPER_EMPLEADOS = True
    Else
        getPER_EMPLEADOS = False
    End If
End Property
Public Property Get getUSO() As String
    getUSO = uso
End Property
Public Property Get getFIRMA() As String
    getFIRMA = firma
End Property
Public Property Get getIMAGEN() As String
    getIMAGEN = imagen
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property
Public Property Get getREVISION_CARGO() As String
    getREVISION_CARGO = REVISION_CARGO
End Property
Public Property Get getRESPONSABLE_DEPARTAMENTOS(ByVal Index As Integer) As Integer
    getRESPONSABLE_DEPARTAMENTOS = RESPONSABLE_DEPARTAMENTOS(Index)
End Property
Public Property Let setRESPONSABLE_DEPARTAMENTOS(ByVal Index As Integer, ByVal Value As Integer)
    RESPONSABLE_DEPARTAMENTOS(Index) = Value
End Property
Public Property Let setRESPONSABLE_DEPARTAMENTOS_INICIAL(ByVal Index As Integer, ByVal Value As Integer)
    RESPONSABLE_DEPARTAMENTOS_INICIAL(Index) = Value
End Property
Public Property Get getDEPARTAMENTOS_USUARIO() As clsGenericCollection
    Set getDEPARTAMENTOS_USUARIO = DEPARTAMENTOS_USUARIO
End Property
Public Property Set setDEPARTAMENTOS_USUARIO(ByRef Value As clsGenericCollection)
    Set DEPARTAMENTOS_USUARIO = Value
End Property
'E0140-I
Public Property Get getEMAIL() As String
    getEMAIL = EMAIL
End Property
Public Property Get getFNMT_RUTA() As String
    getFNMT_RUTA = FNMT_RUTA
End Property
Public Property Get getFNMT_PASS() As String
    getFNMT_PASS = FNMT_PASS
End Property
'E0140-F
'M1144-I
Public Property Get getPER_TRAMITACION_CONTRATA() As Boolean
    If PER_TRAMITACION_CONTRATA = 1 Then
        getPER_TRAMITACION_CONTRATA = True
    Else
        getPER_TRAMITACION_CONTRATA = False
    End If
End Property

Public Property Get getPER_RFI() As Boolean
    If PER_RFI = 1 Then
        getPER_RFI = True
    Else
        getPER_RFI = False
    End If
End Property

Public Property Get getPER_PFA() As Boolean
    If PER_PFA = 1 Then
        getPER_PFA = True
    Else
        getPER_PFA = False
    End If
End Property
'M1144-F
Public Property Get getPER_SCG() As Boolean
    If PER_SCG = 1 Then
        getPER_SCG = True
    Else
        getPER_SCG = False
    End If
End Property
Public Property Get getPER_TESORERIA_MENU() As Boolean
    If PER_TESORERIA_MENU = 1 Then
        getPER_TESORERIA_MENU = True
    Else
        getPER_TESORERIA_MENU = False
    End If
End Property
Public Property Get getPER_TESORERIA_FP() As Boolean
    If PER_TESORERIA_FP = 1 Then
        getPER_TESORERIA_FP = True
    Else
        getPER_TESORERIA_FP = False
    End If
End Property
'M1377-I
Public Property Get getPER_NORMAS_NO_CONTROLADAS() As Boolean
    If PER_NORMAS_NO_CONTROLADAS = 1 Then
        getPER_NORMAS_NO_CONTROLADAS = True
    Else
        getPER_NORMAS_NO_CONTROLADAS = False
    End If
End Property
Public Property Get getPER_DES_PRODUCTO() As Boolean
    If PER_DES_PRODUCTO = 1 Then
        getPER_DES_PRODUCTO = True
    Else
        getPER_DES_PRODUCTO = False
    End If
End Property
Public Property Get getPER_MOD_CLIENTE() As Boolean
    If PER_MOD_CLIENTE = 1 Then
        getPER_MOD_CLIENTE = True
    Else
        getPER_MOD_CLIENTE = False
    End If
End Property
Public Property Get getPER_INCIDENCIAS() As Boolean
    If PER_INCIDENCIAS = 1 Then
        getPER_INCIDENCIAS = True
    Else
        getPER_INCIDENCIAS = False
    End If
End Property
Public Property Get getPER_PRODUCTIVIDAD() As Boolean
    If PER_PRODUCTIVIDAD = 1 Then
        getPER_PRODUCTIVIDAD = True
    Else
        getPER_PRODUCTIVIDAD = False
    End If
End Property
'M1377-F
Public Function CARGAR(ByVal dato As Long) As Boolean
    Dim rs As New ADODB.Recordset, rs_dpto As New ADODB.Recordset
    Dim consulta As String, consulta_responsables As String
    Dim objdep As clsGenericClass
    On Error GoTo fallo
    consulta = "SELECT * FROM usuarios WHERE ID_EMPLEADO =" & dato
    'jonathan.2009.10.27 ------------------------------------------------------------------
    ' Consulta los parametros de departamentos para conocer aquellos de los que es responsable
    consulta_responsables = "SELECT * FROM decodificadora WHERE codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & " AND PARAMETROS Like '%idresp=" & dato & ";%'"
    Set rs_dpto = datos_bd(consulta_responsables, True)
    If rs_dpto.RecordCount <> 0 Then
       rs_dpto.MoveFirst
       While Not rs_dpto.EOF
           RESPONSABLE_DEPARTAMENTOS(CInt(rs_dpto("VALOR"))) = vbChecked
           RESPONSABLE_DEPARTAMENTOS_INICIAL(CInt(rs_dpto("VALOR"))) = vbChecked
           rs_dpto.MoveNext
       Wend
    End If
    rs_dpto.Close
    Set rs_dpto = Nothing
    'fin jonathan ------------------------------------------------------------------
    
    'jonathan.2009.11.17 ------------------------------------------------------------------
    ' Consulta departamentos del usuario
    consulta_responsables = "select usuarios_departamentos.ID_USUARIO, usuarios_departamentos.ID_DEPARTAMENTO, d1.DESCRIPCION"
    consulta_responsables = consulta_responsables & " From usuarios_departamentos"
    consulta_responsables = consulta_responsables & " LEFT JOIN   decodificadora as d1"
    consulta_responsables = consulta_responsables & " on usuarios_departamentos.ID_DEPARTAMENTO = d1.valor and d1.CODIGO = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & " WHERE usuarios_departamentos.ID_USUARIO = " & CStr(dato) & ";"

    Set rs_dpto = datos_bd(consulta_responsables, True)
    If rs_dpto.RecordCount <> 0 Then
       rs_dpto.MoveFirst
       While Not rs_dpto.EOF
            Set objdep = New clsGenericClass
            objdep.setID = rs_dpto("ID_DEPARTAMENTO")
            objdep.setDESCRIPCION = rs_dpto("DESCRIPCION")
            Call DEPARTAMENTOS_USUARIO.Add(objdep, objdep.getID)
            rs_dpto.MoveNext
       Wend
    End If
    rs_dpto.Close
    Set rs_dpto = Nothing
    'fin jonathan ------------------------------------------------------------------
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        ID_EMPLEADO = rs("ID_EMPLEADO")
        nombre = rs("NOMBRE")
        APELLIDOS = rs("APELLIDOS")
        USUARIO = rs("USUARIO")
        PassWord = rs("password")
        PER_IMPRESION = rs("PER_IMPRESION")
        PER_FACTURACION = rs("PER_FACTURACION")
        PER_MODIFICACION = rs("PER_MODIFICACION")
        PER_ELIMINACION = rs("PER_ELIMINACION")
        PER_USUARIOS = rs("PER_USUARIOS")
        PER_EDICION = rs("PER_EDICION")
        PER_CIERRE = rs("PER_CIERRE")
        PER_PEDIDOS_REACTIVOS = rs("PER_PEDIDOS_REACTIVOS")
        PER_EMPLEADOS = rs("PER_EMPLEADOS")
        PER_CONTABILIDAD = rs("PER_CONTABILIDAD")
        PER_MATRIZ_CUALIF = rs("PER_MATRIZ_CUALIF")
        PER_ENVIO_PEDIDOS_PROVEEDOR = rs("PER_ENVIO_PEDIDOS_PROVEEDOR")
        PER_DOCUMENTACION_CALIDAD = rs("PER_DOCUMENTACION_CALIDAD")
        PER_PROYECTOS = rs("PER_PROYECTOS")
        PER_NC = rs("PER_NC")
        PER_DATOS_ESPECIALES = rs("PER_DATOS_ESPECIALES")
        PER_PNT = rs("PER_PNT")
        PER_ADMIN_PNT = rs("PER_ADMIN_PNT")
        PER_VIDEOS = CInt(rs("PER_VIDEOS"))
        PER_IMPRESION_PNT = rs("PER_IMPRESION_PNT")
        PER_INCURRIDOS = rs("PER_INCURRIDOS")
        PER_REVISION = rs("PER_REVISION")
        PER_OFERTAS = rs("PER_OFERTAS")
        PER_PLAZO_ENTREGA_CAMBIO = rs("PER_PLAZO_ENTREGA_CAMBIO")
        PER_PLAZO_ENTREGA_LISTADO = rs("PER_PLAZO_ENTREGA_LISTADO")
        PER_INDICADORES_CLIENTE = rs("PER_INDICADORES_CLIENTE")
        PER_FAMILIAS_CA = rs("PER_FAMILIAS_CA")
        
        REVISION_CARGO = rs("REVISION_CARGO")
        uso = rs("USO")
        firma = rs("FIRMA")
        imagen = rs("IMAGEN")
        anulado = rs("ANULADO")
        'E0144-I
        EMAIL = rs("EMAIL")
        FNMT_RUTA = rs("FNMT_RUTA")
        FNMT_PASS = rs("FNMT_PASS")
        'E0144-I
        'M1144-I
        PER_TRAMITACION_CONTRATA = rs("PER_TRAMITACION_CONTRATA")
        PER_RFI = rs("PER_RFI")
        PER_PFA = rs("PER_PFA")
        'M1144-F
        PER_SCG = rs("PER_SCG")
        PER_TESORERIA_MENU = rs("PER_TESORERIA_MENU")
        PER_TESORERIA_FP = rs("PER_TESORERIA_FP")
        'M1377-I
        PER_NORMAS_NO_CONTROLADAS = rs("PER_NORMAS_NO_CONTROLADAS")
        PER_DES_PRODUCTO = rs("PER_DES_PRODUCTO")
        PER_MOD_CLIENTE = rs("PER_MOD_CLIENTE")
        PER_INCIDENCIAS = rs("PER_INCIDENCIAS")
        PER_PRODUCTIVIDAD = rs("PER_PRODUCTIVIDAD")
        'M1377-F
        CARGAR = True
    Else
        CARGAR = False
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
fallo:
    CARGAR = False
    log "Error al recuperar el usuario" & Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM usuarios WHERE ANULADO = 0 AND ID_EMPLEADO <> 0 order by APELLIDOS, NOMBRE"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_Usuario() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM usuarios WHERE ANULADO = 0 AND ID_EMPLEADO <> 0 order by usuario"
    Set Listado_por_Usuario = datos_bd(consulta)
End Function
Public Function Listado_Filtro(lApellido As String, lNombre As String, anulado As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM usuarios " & _
               " WHERE ANULADO = " & anulado & _
               "   AND APELLIDOS LIKE '%" & lApellido & "%' " & _
               "   AND NOMBRE LIKE '%" & lNombre & "%' " & _
               "   AND ID_EMPLEADO <> 0 " & _
               " ORDER BY APELLIDOS, NOMBRE"
    Set Listado_Filtro = datos_bd(consulta)
End Function


Public Function Listado_por_departamento(ByVal prmID_DEPATAMENTO As Long) As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT * FROM usuarios "
    consulta = consulta & " join usuarios_departamentos on usuarios.ID_EMPLEADO = usuarios_departamentos.id_usuario and usuarios_departamentos.id_departamento = " & CStr(prmID_DEPATAMENTO)
    consulta = consulta & " WHERE ANULADO = 0 order by id_empleado"
    
    
    Set Listado_por_departamento = datos_bd(consulta)
End Function


Public Function Listado_responsables_departamento(ByVal prmID_DEPATAMENTO As Long) As String
    Dim consulta As String, res As String
    Dim rs As New ADODB.Recordset
    
    
    consulta = "select parametros from decodificadora where codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & " and valor = " & CStr(prmID_DEPATAMENTO)
    res = ""
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        res = rs!parametros
    End If
    
    Listado_responsables_departamento = res
    
End Function

Public Function Eliminar(EMPLEADO As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE usuarios SET ANULADO = 1 WHERE ID_EMPLEADO = " & EMPLEADO
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar al empleado. " & Err.Number & "/" & Err.Description
End Function
Public Sub CrearID()
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_EMPLEADO) FROM usuarios"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_EMPLEADO = 1
    Else
        ID_EMPLEADO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox Err.Number & "/" & Err.Description
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "select * from usuarios where usuario = '" & USUARIO & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "El usuario ya existe.", vbCritical, App.Title
        Insertar = 0
        Exit Function
    End If
    PER_EMPLEADOS = 0
    Me.CrearID
    'E0141-I
    consulta = "Insert into usuarios " & _
               " values(" & _
               ID_EMPLEADO & ",'" & nombre & "','" & _
               APELLIDOS & "','" & USUARIO & "','" & _
               PassWord & "'," & PER_IMPRESION & "," & _
               PER_FACTURACION & "," & PER_MODIFICACION & "," & _
               PER_ELIMINACION & "," & PER_USUARIOS & "," & _
               PER_EDICION & "," & PER_CIERRE & "," & PER_PEDIDOS_REACTIVOS & "," & PER_EMPLEADOS & "," & _
               PER_CONTABILIDAD & "," & PER_ENVIO_PEDIDOS_PROVEEDOR & "," & _
               PER_DOCUMENTACION_CALIDAD & "," & PER_PROYECTOS & "," & PER_NC & "," & PER_DATOS_ESPECIALES & "," & _
               "'" & uso & "','" & firma & "','" & imagen & "'," & anulado & ",'" & EMAIL & "','" & _
               FNMT_RUTA & "','" & FNMT_PASS & "'," & PER_PNT & "," & PER_ADMIN_PNT & "," & PER_MATRIZ_CUALIF & "," & PER_VIDEOS & "," & PER_IMPRESION_PNT & "," & PER_INCURRIDOS & "," & PER_REVISION & "," & PER_OFERTAS & "," & _
               PER_PLAZO_ENTREGA_CAMBIO & "," & PER_PLAZO_ENTREGA_LISTADO & "," & PER_INDICADORES_CLIENTE & "," & PER_FAMILIAS_CA & ",'" & REVISION_CARGO & "'," & PER_TRAMITACION_CONTRATA & "," & PER_RFI & "," & PER_PFA & "," & PER_SCG & "," & PER_TESORERIA_MENU & "," & PER_TESORERIA_FP & "," & PER_NORMAS_NO_CONTROLADAS & "," & PER_DES_PRODUCTO & "," & PER_MOD_CLIENTE & "," & PER_INCIDENCIAS & "," & PER_PRODUCTIVIDAD & ")"
    'E0141-I
    execute_bd consulta
    
    Call GuaradarResponsableDepartamentos(ID_EMPLEADO)
    Call GuardarDepartamentosUsuario(ID_EMPLEADO)
    Me.SubirFirma ID_EMPLEADO, firma

    Insertar = ID_EMPLEADO
    Exit Function
fallo:
    MsgBox "Error al insertar el empleado (Insertar)", vbCritical, Err.Description
    Insertar = 0
End Function
Public Function Modificar(ID As Integer) As Long
    On Error GoTo fallo
    Dim consulta As String
    'E0142-I
    ' SOLO SE PUEDEN ASIGNAR POR BASE DE DATOS
    ' PER_ADMIN_PNT
    ' PER_EMPLEADOS
    consulta = "UPDATE usuarios SET " & _
               " NOMBRE = '" & nombre & "'," & _
               " APELLIDOS = '" & APELLIDOS & "'," & _
               " USUARIO = '" & USUARIO & "'," & _
               " PASSWORD = '" & PassWord & "'," & _
               " PER_IMPRESION = " & PER_IMPRESION & "," & _
               " PER_FACTURACION = " & PER_FACTURACION & "," & _
               " PER_MODIFICACION = " & PER_MODIFICACION & ","
    
    consulta = consulta & " PER_ELIMINACION = " & PER_ELIMINACION & "," & _
               " PER_EDICION = " & PER_EDICION & "," & _
               " PER_CIERRE = " & PER_CIERRE & "," & _
               " PER_PEDIDOS_REACTIVOS = " & PER_PEDIDOS_REACTIVOS & "," & _
               " PER_CONTABILIDAD = " & PER_CONTABILIDAD & "," & _
               " PER_MATRIZ_CUALIF = " & PER_MATRIZ_CUALIF & "," & _
               " PER_ENVIO_PEDIDOS_PROVEEDOR = " & PER_ENVIO_PEDIDOS_PROVEEDOR & "," & _
               " PER_DOCUMENTACION_CALIDAD = " & PER_DOCUMENTACION_CALIDAD & "," & _
               " PER_PROYECTOS = " & PER_PROYECTOS & "," & _
               " PER_NC = " & PER_NC & "," & _
               " PER_DATOS_ESPECIALES = " & PER_DATOS_ESPECIALES & "," & _
               " PER_PNT = " & PER_PNT & "," & _
               " USO = '" & uso & "'," & _
               " FIRMA = '" & firma & "'," & _
               " IMAGEN = '" & imagen & "',"
    consulta = consulta & " PER_USUARIOS = " & PER_USUARIOS & ", " & _
               " PER_IMPRESION_PNT = " & PER_IMPRESION_PNT & ", " & _
               " PER_INCURRIDOS = " & PER_INCURRIDOS & ", " & _
               " PER_REVISION = " & PER_REVISION & ", " & _
               " PER_OFERTAS = " & PER_OFERTAS & ", " & _
               " PER_PLAZO_ENTREGA_CAMBIO = " & PER_PLAZO_ENTREGA_CAMBIO & ", " & _
               " PER_PLAZO_ENTREGA_LISTADO = " & PER_PLAZO_ENTREGA_LISTADO & ", " & _
               " PER_INDICADORES_CLIENTE = " & PER_INDICADORES_CLIENTE & ", " & _
               " PER_FAMILIAS_CA = " & PER_FAMILIAS_CA & ", " & _
               " REVISION_CARGO = '" & REVISION_CARGO & "', " & _
               " PER_VIDEOS = " & PER_VIDEOS & ", " & _
               " EMAIL = '" & EMAIL & "'," & _
               " FNMT_RUTA = '" & FNMT_RUTA & "',FNMT_PASS='" & FNMT_PASS & "'," & _
               " PER_TRAMITACION_CONTRATA = " & PER_TRAMITACION_CONTRATA & ", " & _
               " PER_RFI = " & PER_RFI & ", " & _
               " PER_PFA = " & PER_PFA & ", " & _
               " PER_SCG = " & PER_SCG & ", " & _
               " PER_TESORERIA_FP = " & PER_TESORERIA_FP & ", " & _
               " PER_TESORERIA_MENU = " & PER_TESORERIA_MENU & ", " & _
               " PER_NORMAS_NO_CONTROLADAS = " & PER_NORMAS_NO_CONTROLADAS & _
               " ,PER_DES_PRODUCTO = " & PER_DES_PRODUCTO & ",PER_MOD_CLIENTE = " & PER_MOD_CLIENTE & _
               " ,PER_INCIDENCIAS = " & PER_INCIDENCIAS & ",PER_PRODUCTIVIDAD = " & PER_PRODUCTIVIDAD & _
               " WHERE ID_EMPLEADO = " & ID
    'E0142-F
    execute_bd consulta
    'jonathan.2009.10.27
    Call GuaradarResponsableDepartamentos(ID)
    Call GuardarDepartamentosUsuario(ID)
    'final.jonathan.2009.10.27
    Me.SubirFirma CLng(ID), firma
    Modificar = ID
    Exit Function
fallo:
    MsgBox "Error al modificar el empleado (Modificar)", vbCritical, Err.Description
    Modificar = 0
End Function
Public Function SubirFirma(EMPLEADO_ID As Long, fichero As String) As String
   On Error GoTo SubirFirma_Error

    execute_bd "DELETE from usuarios_firmas where ID_EMPLEADO = " & EMPLEADO_ID
    If fichero = "" Then Exit Function
    If Dir(fichero) = "" Then Exit Function
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    
    Dim conn As ADODB.Connection
    CrearConexionGlobal conn, "", ""
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM usuarios_firmas WHERE 1=0", conn, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID_EMPLEADO = EMPLEADO_ID
    rs!DESCRIPCION = fichero
    rs!informe = ""
    rs!CARPETA = ""
    rs!TAMANO = ""
    rs!firma = mystream.Read
    rs.Update
    rs.Close
    mystream.Close

   On Error GoTo 0
   Exit Function

SubirFirma_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SubirFirma of Módulo de clase clsUsuarios"
End Function

Public Function deslogonear(ID As Integer) As Long
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE usuarios SET " & _
               " USO = ''" & _
               " WHERE ID_EMPLEADO = " & ID
    execute_bd consulta
    deslogonear = ID
    Exit Function
fallo:
    MsgBox "Error al deslogonear al empleado (Modificar)", vbCritical, Err.Description
    deslogonear = 0
End Function

Private Function GuaradarResponsableDepartamentos(ByVal idEmpleado As Integer)
Dim intCont As Integer, intMax As Integer
Dim strSql() As String
ReDim Preserve strSql(0)
intMax = 0

    For intCont = 1 To 10
        ' SI ANTES ERA RESPONSABLE Y AHORA NO
        If RESPONSABLE_DEPARTAMENTOS_INICIAL(intCont) = vbChecked And RESPONSABLE_DEPARTAMENTOS(intCont) = vbUnchecked Then
            intMax = intMax + 1
            ReDim Preserve strSql(intMax)
            strSql(intMax) = "UPDATE decodificadora SET PARAMETROS=REPLACE(PARAMETROS,'idresp=" & CStr(idEmpleado) & ";','') WHERE CODIGO=107 AND VALOR=" & intCont
        ElseIf RESPONSABLE_DEPARTAMENTOS_INICIAL(intCont) = vbUnchecked And RESPONSABLE_DEPARTAMENTOS(intCont) = vbChecked Then
            intMax = intMax + 1
            ReDim Preserve strSql(intMax)
            strSql(intMax) = "UPDATE decodificadora SET PARAMETROS='idresp=" & CStr(idEmpleado) & ";' WHERE CODIGO=107 AND VALOR=" & CStr(intCont) & " AND PARAMETROS = 'idresp=0;'"
            intMax = intMax + 1
            ReDim Preserve strSql(intMax)
            strSql(intMax) = "UPDATE decodificadora SET PARAMETROS=concat(PARAMETROS,'idresp=" & CStr(idEmpleado) & ";') WHERE CODIGO=107 AND VALOR=" & CStr(intCont) & " AND PARAMETROS != 'idresp=0;' AND PARAMETROS NOT LIKE '%idresp=" & CStr(idEmpleado) & ";%'"
        End If
    Next intCont
    
    
    For intCont = 1 To intMax
        If Trim(strSql(intCont)) <> "" Then
            execute_bd strSql(intCont)
            'Debug.Print strSQL(intCont)
        End If
    Next intCont

End Function

Public Function Listado_Logo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT nombre,uso FROM usuarios where uso <> '' order by id_empleado"
    Set Listado_Logo = datos_bd(consulta)
End Function
Public Function Autentificacion(user As String, PASS As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta_responsables As String
    Dim rs_dpto As New ADODB.Recordset
        
    Dim consulta As String
    On Error GoTo fallo
    consulta = "SELECT * FROM usuarios " & _
               " WHERE usuario='" & user & "'" & _
               "   AND password='" & Encripta(PASS, user) & "'" & _
               "   AND anulado=0"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Autentificacion = False
    Else
'        If Trim(rs("uso")) <> "" Then
'           MsgBox "El usuario ya esta en uso en la máquina " & rs("uso") & ".", vbExclamation, App.Title
'        End If
        ID_EMPLEADO = rs("ID_EMPLEADO")
        nombre = rs("NOMBRE")
        APELLIDOS = rs("APELLIDOS")
        USUARIO = rs("USUARIO")
        PER_IMPRESION = rs("PER_IMPRESION")
        PER_FACTURACION = rs("PER_FACTURACION")
        PER_MODIFICACION = rs("PER_MODIFICACION")
        PER_ELIMINACION = rs("PER_ELIMINACION")
        PER_USUARIOS = rs("PER_USUARIOS")
        PER_EDICION = rs("PER_EDICION")
        PER_CIERRE = rs("PER_CIERRE")
        PER_PEDIDOS_REACTIVOS = rs("PER_PEDIDOS_REACTIVOS")
        PER_EMPLEADOS = rs("PER_EMPLEADOS")
        PER_CONTABILIDAD = rs("PER_CONTABILIDAD")
        PER_MATRIZ_CUALIF = rs("PER_MATRIZ_CUALIF")
        PER_ENVIO_PEDIDOS_PROVEEDOR = rs("PER_ENVIO_PEDIDOS_PROVEEDOR")
        PER_DOCUMENTACION_CALIDAD = rs("PER_DOCUMENTACION_CALIDAD")
        PER_PROYECTOS = rs("PER_PROYECTOS")
        PER_NC = rs("PER_NC")
        PER_DATOS_ESPECIALES = rs("PER_DATOS_ESPECIALES")
        PER_PNT = rs("PER_PNT")
        PER_ADMIN_PNT = rs("PER_ADMIN_PNT")
        PER_VIDEOS = CInt(rs("PER_VIDEOS"))
        PER_IMPRESION_PNT = rs("PER_IMPRESION_PNT")
        PER_INCURRIDOS = rs("PER_INCURRIDOS")
        PER_REVISION = rs("PER_REVISION")
        PER_OFERTAS = rs("PER_OFERTAS")
        PER_PLAZO_ENTREGA_CAMBIO = rs("PER_PLAZO_ENTREGA_CAMBIO")
        PER_PLAZO_ENTREGA_LISTADO = rs("PER_PLAZO_ENTREGA_LISTADO")
        PER_INDICADORES_CLIENTE = rs("PER_INDICADORES_CLIENTE")
        PER_FAMILIAS_CA = rs("PER_FAMILIAS_CA")
        REVISION_CARGO = rs("REVISION_CARGO")
        uso = rs("USO")
        firma = rs("FIRMA")
        imagen = rs("IMAGEN")
        EMAIL = rs("EMAIL")
        FNMT_RUTA = rs("FNMT_RUTA")
        FNMT_PASS = rs("FNMT_PASS")
        'M1144-I
        PER_TRAMITACION_CONTRATA = rs("PER_TRAMITACION_CONTRATA")
        PER_RFI = rs("PER_RFI")
        PER_PFA = rs("PER_PFA")
        'M1144-F
        PER_SCG = rs("PER_SCG")
        PER_TESORERIA_MENU = rs("PER_TESORERIA_MENU")
        PER_TESORERIA_FP = rs("PER_TESORERIA_FP")
        'M1377-I
        PER_NORMAS_NO_CONTROLADAS = rs("PER_NORMAS_NO_CONTROLADAS")
        PER_DES_PRODUCTO = rs("PER_DES_PRODUCTO")
        PER_MOD_CLIENTE = rs("PER_MOD_CLIENTE")
        PER_INCIDENCIAS = rs("PER_INCIDENCIAS")
        PER_PRODUCTIVIDAD = rs("PER_PRODUCTIVIDAD")
        'M1377-F
       'jonathan.2009.10.27 ------------------------------------------------------------------
        ' Consulta los parametros de departamentos para conocer aquellos de los que es responsable
        consulta_responsables = "SELECT * FROM decodificadora WHERE codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & " AND PARAMETROS like '%idresp=" & ID_EMPLEADO & ";%'"
        Set rs_dpto = datos_bd(consulta_responsables, True)
        If rs_dpto.RecordCount <> 0 Then
           rs_dpto.MoveFirst
           While Not rs_dpto.EOF
               RESPONSABLE_DEPARTAMENTOS(CInt(rs_dpto("VALOR"))) = vbChecked
               RESPONSABLE_DEPARTAMENTOS_INICIAL(CInt(rs_dpto("VALOR"))) = vbChecked
               rs_dpto.MoveNext
           Wend
        End If
        rs_dpto.Close
        Set rs_dpto = Nothing
       'fin jonathan ------------------------------------------------------------------
        
        
        
        Autentificacion = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
fallo:
    MsgBox "Error al autentificar el usuario. " & Err.Description, vbCritical, App.Title
    Autentificacion = False
End Function
Public Function modificar_password(ID As Integer, PASS As String) As Long
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE usuarios SET " & _
               " PASSWORD = '" & PASS & "'" & _
               " WHERE ID_EMPLEADO = " & ID
    execute_bd consulta
    modificar_password = ID
    Exit Function
fallo:
    modificar_password = 0
    MsgBox "Error al modificar el password del empleado (Modificar)", vbCritical, Err.Description
End Function

Public Function firma_director() As String
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT FIRMA FROM usuarios WHERE ID_EMPLEADO = 1"
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        firma_director = rs(0)
    End If
    Exit Function
fallo:
    log ("Error al recuperar la firma del director.")
End Function

Public Function firma_analista(MUESTRA As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (MUESTRA)
    If oMuestra.getCERRADA <> 0 Then
'          Dim oDeter As New clsDeterminaciones
'          Dim rs As ADODB.Recordset
'          If oMuestra.getANALISIS_MODIFICADO = 3 Then
'              Set rs = datos_bd("SELECT EMPLEADO_ID FROM SELLANTES_RESULTADOS WHERE MUESTRA_ID = " & MUESTRA & " ORDER BY ORDEN")
'          Else
'              Set rs = oDeter.lista_determinaciones(MUESTRA)
'          End If
'          If rs.RecordCount <> 0 Then
'          oMuestra.getCERRADA_USUARIO
'             firma_analista = ReadINI(App.Path + "\config.ini", "documentos", "firmas") & "\" & rs("EMPLEADO_ID") & ".jpg"
'          End If
          firma_analista = ReadINI(App.Path + "\config.ini", "documentos", "firmas") & "\" & oMuestra.getCERRADA_USUARIO & ".jpg"
    
    End If
    Exit Function
fallo:
    log ("Error al recuperar la firma del analista.")
End Function
Public Function firma_responsable_cierre(MUESTRA As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    firma_responsable_cierre = ""
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (MUESTRA)
    If oMuestra.getCERRADA = 1 Then
          consulta = "SELECT FIRMA FROM usuarios WHERE ID_EMPLEADO = " & oMuestra.getCERRADA_USUARIO
          Dim rs As ADODB.Recordset
          Set rs = datos_bd(consulta)
          If rs.RecordCount <> 0 Then
             firma_responsable_cierre = rs(0)
          End If
    End If
    Exit Function
fallo:
    log ("Error al recuperar la firma del responsable de cierre.")
End Function

Public Function nombre_analista(MUESTRA As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (MUESTRA)
    If oMuestra.getCERRADA <> 0 Then
          Dim oDeter As New clsDeterminaciones
          Dim rs As ADODB.Recordset
          If oMuestra.getANALISIS_MODIFICADO = 3 Then
              Set rs = datos_bd("SELECT EMPLEADO_ID FROM sellantes_resultados WHERE MUESTRA_ID = " & MUESTRA & " ORDER BY ORDEN")
          Else
              Set rs = oDeter.lista_determinaciones(MUESTRA)
          End If
          If rs.RecordCount <> 0 Then
             consulta = "SELECT concat(NOMBRE,' ',APELLIDOS) FROM usuarios WHERE ID_EMPLEADO = " & rs("EMPLEADO_ID")
             Set rs = datos_bd(consulta)
             If rs.RecordCount <> 0 Then
                nombre_analista = rs(0)
             End If
          End If
    End If
    Exit Function
fallo:
    MsgBox "Error al recuperar la firma del analista", vbCritical, Err.Description
End Function
Public Function firma_recepcion(MUESTRA As Long) As String
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (MUESTRA)
    consulta = "SELECT FIRMA FROM usuarios WHERE ID_EMPLEADO = " & oMuestra.getEMPLEADO_ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        firma_recepcion = rs(0)
    End If
    Exit Function
fallo:
    MsgBox "Error al recuperar la firma del recepcionista", vbCritical, Err.Description
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_EMPLEADO, concat(NOMBRE,' ',APELLIDOS) FROM usuarios " & _
               " WHERE ANULADO = 0 " & _
               " ORDER BY 2"
    Set Listado_Combo = datos_bd(consulta)
End Function

'E0054-I
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "usuarios"
        .setDESCRIPCION = "Usuarios"
        .setPK = "ID_EMPLEADO"
        .setCAMPO = "concat(NOMBRE,' ',APELLIDOS)"
'        .setMUESTRA_DETALLE = False
        .setMUESTRA_DETALLE = True
        .setFILTRO = "ANULADO=0 " & filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function
'E0054-F
Private Sub Class_Initialize()
    DEPARTAMENTOS_USUARIO.KeyName = "setID"
End Sub

Private Sub GuardarDepartamentosUsuario(ByVal ID_USUARIO As Long)

    Dim objdep As clsGenericClass
    Dim strSql() As String
    Dim intCont As Long
    intCont = 0
    
    For Each objdep In DEPARTAMENTOS_USUARIO.Iterator
    
        If objdep.getID_AUX = enumIdAux.ID_AUX_ELIMINADO Then
            intCont = intCont + 1
            ReDim Preserve strSql(intCont)
            strSql(intCont - 1) = "DELETE FROM usuarios_departamentos WHERE ID_USUARIO = " & CStr(ID_USUARIO) & " AND ID_DEPARTAMENTO = " & objdep.getID & "; "
        ElseIf objdep.getID_AUX = enumIdAux.ID_AUX_NUEVO Then
            intCont = intCont + 1
            ReDim Preserve strSql(intCont)
            strSql(intCont - 1) = "INSERT INTO usuarios_departamentos (ID_USUARIO,ID_DEPARTAMENTO) VALUES(" & CStr(ID_USUARIO) & ", " & objdep.getID & "); "
        End If
    
    Next objdep
    
    If intCont > 0 Then
        For intCont = 0 To (UBound(strSql) - 1)
            execute_bd strSql(intCont)
        Next intCont
    End If
    

End Sub
