VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOBRAS
'***************************************************************
Private ID_OBRA As Long
Private CLIENTE_ID As Long
Private NOMBRE As String
Private DIRECCION As String
Private CP As String
Private MUNICIPIO_ID As Long
Private PROVINCIA_ID As Long
Private Telefono As String
Private FORMA_PAGO_ID As Long
Private descuento As String
Private contrapartida As Long
Private EMAIL As String
Private BANCO As String
Private ccc As String
Private FECHA_OFERTA As String
Private TARIFA_PORTE_ID As Long
Private TIPO_FACTURACION As String
Private COMERCIAL_ID As Long
Private BANCO_DIRECCION As String
Private OBSERVACIONES As String
Private AVISOS As String
Private REFERENCIA_FACTURA As String
Private DESTINO_FACTURA As String
Private TIPO_OBRA_ID As Integer
Private TIPO_IVA As Integer
Private libro As Integer
Private TIPO_IMPORTE_ID As Integer

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOBRAS
'***************************************************************
Public Property Let setID_OBRA(ByVal dato As Long)
        ID_OBRA = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDIRECCION(ByVal dato As String)
        DIRECCION = dato
End Property
Public Property Let setCP(ByVal dato As String)
        CP = dato
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
        MUNICIPIO_ID = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
        PROVINCIA_ID = dato
End Property
Public Property Let setTelefono(ByVal dato As String)
        Telefono = dato
End Property
Public Property Let setFORMA_PAGO_ID(ByVal dato As Long)
        FORMA_PAGO_ID = dato
End Property
Public Property Let setDESCUENTO(ByVal dato As String)
        descuento = dato
End Property
Public Property Let setCONTRAPARTIDA(ByVal dato As Long)
        contrapartida = dato
End Property
Public Property Let setEMAIL(ByVal dato As String)
        EMAIL = dato
End Property
Public Property Let setBANCO(ByVal dato As String)
        BANCO = dato
End Property
Public Property Let setCCC(ByVal dato As String)
        ccc = dato
End Property
Public Property Let setFECHA_OFERTA(ByVal dato As String)
        FECHA_OFERTA = dato
End Property
Public Property Let setTARIFA_PORTE_ID(ByVal dato As Long)
        TARIFA_PORTE_ID = dato
End Property
Public Property Let setTIPO_FACTURACION(ByVal dato As String)
        TIPO_FACTURACION = dato
End Property
Public Property Let setCOMERCIAL_ID(ByVal dato As Long)
        COMERCIAL_ID = dato
End Property
Public Property Let setBANCO_DIRECCION(ByVal dato As String)
        BANCO_DIRECCION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setAVISOS(ByVal dato As String)
    AVISOS = Trim(dato)
End Property
Public Property Let setREFERENCIA_FACTURA(ByVal dato As String)
    REFERENCIA_FACTURA = Trim(dato)
End Property
Public Property Let setDESTINO_FACTURA(ByVal dato As String)
    DESTINO_FACTURA = Trim(dato)
End Property
Public Property Let setTIPO_OBRA_ID(ByVal dato As Integer)
    TIPO_OBRA_ID = Trim(dato)
End Property
Public Property Let setTIPO_IVA(ByVal dato As Integer)
    TIPO_IVA = Trim(dato)
End Property
Public Property Let setLIBRO(ByVal dato As Integer)
    libro = Trim(dato)
End Property
Public Property Let setTIPO_IMPORTE_ID(ByVal dato As Integer)
    TIPO_IMPORTE_ID = Trim(dato)
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOBRAS
'***************************************************************
Public Property Get getID_OBRA() As Long
        getID_OBRA = ID_OBRA
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDIRECCION() As String
        getDIRECCION = DIRECCION
End Property
Public Property Get getCP() As String
        getCP = CP
End Property
Public Property Get getMUNICIPIO_ID() As Long
        getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getPROVINCIA_ID() As Long
        getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getTelefono() As String
        getTelefono = Telefono
End Property
Public Property Get getFORMA_PAGO_ID() As Long
        getFORMA_PAGO_ID = FORMA_PAGO_ID
End Property
Public Property Get getDESCUENTO() As String
        getDESCUENTO = descuento
End Property
Public Property Get getCONTRAPARTIDA() As Long
        getCONTRAPARTIDA = contrapartida
End Property
Public Property Get getEMAIL() As String
        getEMAIL = EMAIL
End Property
Public Property Get getBANCO() As String
        getBANCO = BANCO
End Property
Public Property Get getCCC() As String
        getCCC = ccc
End Property
Public Property Get getFECHA_OFERTA() As String
        getFECHA_OFERTA = FECHA_OFERTA
End Property
Public Property Get getTARIFA_PORTE_ID() As Long
        getTARIFA_PORTE_ID = TARIFA_PORTE_ID
End Property
Public Property Get getTIPO_FACTURACION() As String
        getTIPO_FACTURACION = TIPO_FACTURACION
End Property
Public Property Get getCOMERCIAL_ID() As Long
        getCOMERCIAL_ID = COMERCIAL_ID
End Property
Public Property Get getBANCO_DIRECCION() As String
        getBANCO_DIRECCION = BANCO_DIRECCION
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getAVISOS() As String
    getAVISOS = AVISOS
End Property
Public Property Get getREFERENCIA_FACTURA() As String
    getREFERENCIA_FACTURA = REFERENCIA_FACTURA
End Property
Public Property Get getDESTINO_FACTURA() As String
    getDESTINO_FACTURA = DESTINO_FACTURA
End Property
Public Property Get getTIPO_OBRA_ID() As Integer
    getTIPO_OBRA_ID = TIPO_OBRA_ID
End Property
Public Property Get getTIPO_IVA() As Integer
    getTIPO_IVA = TIPO_IVA
End Property
Public Property Get getLIBRO() As Integer
    getLIBRO = libro
End Property
Public Property Get getTIPO_IMPORTE_ID() As Integer
    getTIPO_IMPORTE_ID = TIPO_IMPORTE_ID
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOBRAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS WHERE ID_OBRA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_OBRA = rs("ID_OBRA")
                CLIENTE_ID = rs("CLIENTE_ID")
                NOMBRE = rs("NOMBRE")
                DIRECCION = rs("DIRECCION")
                CP = rs("CP")
                MUNICIPIO_ID = rs("MUNICIPIO_ID")
                PROVINCIA_ID = rs("PROVINCIA_ID")
                Telefono = rs("TELEFONO")
                FORMA_PAGO_ID = rs("FORMA_PAGO_ID")
                descuento = rs("DESCUENTO")
                contrapartida = rs("CONTRAPARTIDA")
                EMAIL = rs("EMAIL")
                BANCO = rs("BANCO")
                ccc = rs("CCC")
                If Not IsNull(rs("FECHA_OFERTA")) Then
                    FECHA_OFERTA = rs("FECHA_OFERTA")
                End If
                TARIFA_PORTE_ID = rs("TARIFA_PORTE_ID")
                TIPO_FACTURACION = rs("TIPO_FACTURACION")
                COMERCIAL_ID = rs("COMERCIAL_ID")
                BANCO_DIRECCION = rs("BANCO_DIRECCION")
                OBSERVACIONES = rs("OBSERVACIONES")
                AVISOS = rs("AVISOS")
                REFERENCIA_FACTURA = rs("REFERENCIA_FACTURA")
                DESTINO_FACTURA = rs("DESTINO_FACTURA")
                
                TIPO_OBRA_ID = rs("TIPO_OBRA_ID")
                TIPO_IVA = rs("TIPO_IVA")
                libro = rs("LIBRO")
                TIPO_IMPORTE_ID = rs("TIPO_IMPORTE_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsObras)", vbCritical, Err.Description
End Function
Public Function CrearID(cliente As Long)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        Dim inicio As String
        Dim fin As String
        inicio = CStr(cliente) & "01"
        fin = CStr(cliente) & "99"
        consulta = "SELECT MAX(ID_OBRA) FROM OBRAS WHERE ID_OBRA BETWEEN " & inicio & " AND " & fin
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_OBRA = CLng(CStr(cliente) & "01")
        Else
                ID_OBRA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_OBRA = 0 Then
            Me.CrearID (CLIENTE_ID)
        End If
        If FECHA_OFERTA = "" Then
            FECHA_OFERTA = "0000-00-00"
        End If
        consulta = "INSERT INTO OBRAS " & _
                   "  VALUES (" & _
                        ID_OBRA & "," & CLIENTE_ID & "," & "'" & NOMBRE & "'" & "," & _
                        "'" & DIRECCION & "'" & "," & "'" & CP & "'" & "," & MUNICIPIO_ID & "," & _
                        PROVINCIA_ID & "," & "'" & Telefono & "'" & "," & FORMA_PAGO_ID & ",'" & descuento & "'," & _
                        contrapartida & "," & "'" & EMAIL & "'" & "," & _
                        "'" & BANCO & "'" & "," & "'" & ccc & "','" & FECHA_OFERTA & "'," & TARIFA_PORTE_ID & ",'" & TIPO_FACTURACION & "'," & COMERCIAL_ID & ",'" & BANCO_DIRECCION & "','" & _
                        REFERENCIA_FACTURA & "','" & DESTINO_FACTURA & "','" & OBSERVACIONES & "','" & AVISOS & "'," & TIPO_OBRA_ID & "," & TIPO_IVA & "," & libro & "," & TIPO_IMPORTE_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_OBRA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsObras)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS SET " & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DIRECCION = '" & DIRECCION & "'," & _
                        " CP = '" & CP & "'," & _
                        " MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
                        " PROVINCIA_ID = " & PROVINCIA_ID & "," & _
                        " TELEFONO = '" & Telefono & "'," & _
                        " FORMA_PAGO_ID = " & FORMA_PAGO_ID & "," & _
                        " DESCUENTO = '" & descuento & "'," & _
                        " CONTRAPARTIDA = " & contrapartida & "," & _
                        " EMAIL = '" & EMAIL & "'," & _
                        " BANCO = '" & BANCO & "'," & _
                        " CCC = '" & ccc & "'," & _
                        " TARIFA_PORTE_ID = " & TARIFA_PORTE_ID & "," & _
                        " TIPO_FACTURACION = '" & TIPO_FACTURACION & "'," & _
                        " COMERCIAL_ID = " & COMERCIAL_ID & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " REFERENCIA_FACTURA = '" & REFERENCIA_FACTURA & "'," & _
                        " DESTINO_FACTURA = '" & DESTINO_FACTURA & "'," & _
                        " AVISOS = '" & AVISOS & "'," & _
                        " TIPO_OBRA_ID = " & TIPO_OBRA_ID & "," & " TIPO_IVA = " & TIPO_IVA & "," & _
                        " LIBRO = " & libro & "," & " TIPO_IMPORTE_ID = " & TIPO_IMPORTE_ID & "," & _
                        " BANCO_DIRECCION = '" & BANCO_DIRECCION & "'" & _
                " WHERE ID_OBRA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsObras)", vbCritical, Err.Description
End Function
Public Function ModificarFOferta(ID As Long, fecha) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS SET " & _
                   " FECHA_OFERTA = '" & fecha & "'" & _
                   " WHERE ID_OBRA = " & ID
        execute_bd consulta
        ModificarFOferta = True
        Exit Function
fallo:
        ModificarFOferta = False
        MsgBox "Error al ModificarFOferta (clsObras)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS WHERE OBRA_ID = " & ID
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar la obra. Tiene albaranes o facturas relacionados.", vbExclamation, App.Title
            Eliminar = False
            Exit Function
        End If
        execute_bd "DELETE FROM OBRAS WHERE ID_OBRA = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsObras)", vbCritical, Err.Description
End Function
Public Function Listado(lCODIGO As String, lOBRA As String, cliente As Long, tipo As String) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        If Trim(lCODIGO) <> "" Then
            s = s & " AND O.ID_OBRA LIKE '%" & lCODIGO & "%' "
        End If
        If Trim(lOBRA) <> "" Then
            s = s & " AND O.NOMBRE LIKE '%" & lOBRA & "%' "
        End If
        If cliente <> 0 Then
            s = s & " AND O.CLIENTE_ID = " & cliente
        End If
        If tipo <> "" Then
            If IsNumeric(tipo) Then
                s = s & " AND O.TIPO_OBRA_ID= " & tipo
            End If
        End If
        consulta = "SELECT O.ID_OBRA,O.NOMBRE,C.NOMBRE,O.DIRECCION,O.TELEFONO " & _
                   "  FROM OBRAS O, CLIENTES C " & _
                   " WHERE O.CLIENTE_ID = C.ID_CLIENTE " & _
                   s & _
                   " ORDER BY O.NOMBRE, O.ID_OBRA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_OBRA,CLIENTE_ID FROM OBRAS ORDER BY CLIENTE_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(combo As miComboBCA, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "OBRAS"
        .setDESCRIPCION = "Obras"
        .setPK = "ID_OBRA"
        .setCAMPO = "CONCAT(cast(ID_OBRA as char),' - ',NOMBRE)"
        .setMUESTRA_DETALLE = True
        .setFILTRO = FILTRO
        Set .FORMULARIO = FORMULARIO
    End With
End Function

Public Function Listado_Comercial() As ADODB.Recordset
    ' LP005
    Dim consulta As String
    consulta = "SELECT A.ID_OBRA,A.COMERCIAL_ID,A.NOMBRE,B.NOMBRE " & _
               "  FROM OBRAS A LEFT JOIN PROVINCIAS B " & _
               " ON A.PROVINCIA_ID = B.ID_PROVINCIA " & _
               " ORDER BY A.NOMBRE"
    Set Listado_Comercial = datos_bd(consulta)
End Function


Public Function modificar_comercial(ID As Long, COMERCIAL As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "UPDATE OBRAS SET" & _
               " COMERCIAL_ID = " & COMERCIAL & _
               " WHERE " & _
               " ID_OBRA = " & ID
    execute_bd consulta
    modificar_comercial = True
    Exit Function
fallo:
    modificar_comercial = False
    MsgBox "Error al modificar el comercial en la bd : " & Err.Description, vbCritical, Err.Number
End Function


