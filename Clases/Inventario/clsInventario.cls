VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSINVENTARIO
'***************************************************************
Private ID As Long
Private TIPO_ID As Long
Private NUMERO As Long
Private nombre As String
Private ip As String
Private GATEWAY As String
Private USUARIO_ID As Long
Private CENTRO_ID As Long
Private TOMA_RED As String
Private ZONA_ID As Long
Private SERIE As String
Private USUARIO As String
Private PASS As String
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSINVENTARIO
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
        NUMERO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setIP(ByVal dato As String)
        ip = dato
End Property
Public Property Let setGATEWAY(ByVal dato As String)
        GATEWAY = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
        CENTRO_ID = dato
End Property
Public Property Let setTOMA_RED(ByVal dato As String)
        TOMA_RED = dato
End Property
Public Property Let setZONA_ID(ByVal dato As Long)
        ZONA_ID = dato
End Property
Public Property Let setSERIE(ByVal dato As String)
        SERIE = dato
End Property
Public Property Let setUSUARIO(ByVal dato As String)
        USUARIO = dato
End Property
Public Property Let setPASS(ByVal dato As String)
        PASS = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSINVENTARIO
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getNUMERO() As Long
        getNUMERO = NUMERO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getIP() As String
        getIP = ip
End Property
Public Property Get getGATEWAY() As String
        getGATEWAY = GATEWAY
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getCENTRO_ID() As Long
        getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getTOMA_RED() As String
        getTOMA_RED = TOMA_RED
End Property
Public Property Get getZONA_ID() As Long
        getZONA_ID = ZONA_ID
End Property
Public Property Get getSERIE() As String
        getSERIE = SERIE
End Property
Public Property Get getUSUARIO() As String
        getUSUARIO = USUARIO
End Property
Public Property Get getPASS() As String
        getPASS = PASS
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSINVENTARIO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM INVENTARIO WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                TIPO_ID = rs("TIPO_ID")
                NUMERO = rs("NUMERO")
                nombre = rs("NOMBRE")
                ip = rs("IP")
                GATEWAY = rs("GATEWAY")
                USUARIO_ID = rs("USUARIO_ID")
                CENTRO_ID = rs("CENTRO_ID")
                TOMA_RED = rs("TOMA_RED")
                ZONA_ID = rs("ZONA_ID")
                SERIE = rs("SERIE")
                USUARIO = rs("USUARIO")
                PASS = rs("PASS")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsInventario)", vbCritical, Err.Description
End Function
Public Function calcularCodigo(TIPO_ID As Long) As String
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        Dim oDecodificadora As New clsDecodificadora
        oDecodificadora.Carga_valor DECODIFICADORA.DECODIFICADORA_INVENTARIO_TIPOS, TIPO_ID
        consulta = "SELECT MAX(NUMERO) FROM INVENTARIO WHERE TIPO_ID = " & TIPO_ID
        Set rs = datos_bd(consulta)
        Dim NUMERO As Integer
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                NUMERO = 1
        Else
                NUMERO = rs.Fields(0) + 1
        End If
        calcularCodigo = oDecodificadora.getPARAMETROS & Format(NUMERO, "0000")
        Set rs = Nothing
        Exit Function
fallo:
        calcularCodigo = ""
        MsgBox "Error al cargar los datos(clsInventario)", vbCritical, Err.Description
End Function
Public Function calcularCodigoID(ID As Long) As String
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Me.Carga ID
        Dim consulta As String
        Dim oDecodificadora As New clsDecodificadora
        oDecodificadora.Carga_valor DECODIFICADORA.DECODIFICADORA_INVENTARIO_TIPOS, Me.getTIPO_ID
        calcularCodigoID = oDecodificadora.getPARAMETROS & Format(Me.getNUMERO, "0000")
        Set rs = Nothing
        Exit Function
fallo:
        calcularCodigoID = ""
        MsgBox "Error al cargar los calcularCodigoID(clsInventario)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM INVENTARIO"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CrearNumero(TIPO_ID As Long)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(NUMERO) FROM INVENTARIO WHERE TIPO_ID = " & TIPO_ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                NUMERO = 1
        Else
                NUMERO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Me.CrearNumero (TIPO_ID)
        consulta = "INSERT INTO INVENTARIO " & _
                   "  VALUES (" & _
                        ID & "," & TIPO_ID & "," & NUMERO & "," & _
                        "'" & nombre & "'" & "," & "'" & ip & "'" & "," & "'" & GATEWAY & "'" & "," & _
                        USUARIO_ID & "," & CENTRO_ID & "," & "'" & TOMA_RED & "'" & "," & _
                        ZONA_ID & "," & "'" & SERIE & "'" & "," & "'" & USUARIO & "'" & "," & _
                        "'" & PASS & "'" & "," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsInventario)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE INVENTARIO SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " IP = '" & ip & "'," & _
                        " GATEWAY = '" & GATEWAY & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " CENTRO_ID = " & CENTRO_ID & "," & _
                        " TOMA_RED = '" & TOMA_RED & "'," & _
                        " ZONA_ID = " & ZONA_ID & "," & _
                        " SERIE = '" & SERIE & "'," & _
                        " USUARIO = '" & USUARIO & "'," & _
                        " PASS = '" & PASS & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsInventario)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM INVENTARIO " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM historial_cambios where tipo = " & HC_INVENTARIO & " and identificador = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsInventario)", vbCritical, Err.Description
End Function
Public Function Listado(nombre As String, tipo As String, responsable As String, Centro As String, ip As String, zona As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1 "
        If Trim(nombre) <> "" Then
            filtro = filtro & " and a.NOMBRE like '%" & nombre & "%'"
        End If
        If tipo <> "" Then
            If IsNumeric(tipo) Then
                filtro = filtro & " and a.tipo_id = " & tipo
            End If
        End If
        If responsable <> "" And responsable <> "0" Then
            If IsNumeric(responsable) Then
                filtro = filtro & " and a.usuario_id = " & responsable
            End If
        End If
        If Centro <> "" Then
            If IsNumeric(Centro) Then
                filtro = filtro & " and a.centro_id = " & Centro
            End If
        End If
        If ip <> "" Then
            filtro = filtro & " and a.ip like '%" & ip & "%'"
        End If
        If zona <> "" Then
            If IsNumeric(zona) Then
                filtro = filtro & " and a.zona_id = " & zona
            End If
        End If
        consulta = "SELECT a.ID,d1.DESCRIPCION,concat(d1.parametros,LPAD(a.numero,4,'0')),a.nombre,c.nombre,ifnull(u.nombre,''),IFNULL(d2.descripcion,''),a.IP,a.OBSERVACIONES FROM INVENTARIO a " & _
                   "  LEFT JOIN decodificadora d1 on a.tipo_id = d1.valor and d1.codigo = " & DECODIFICADORA.DECODIFICADORA_INVENTARIO_TIPOS & _
                   "  LEFT JOIN centros c on a.centro_id = c.id_centro " & _
                   "  LEFT JOIN empleados u on a.usuario_id = u.id_empleado " & _
                   "  LEFT JOIN decodificadora d2 on a.zona_id = d2.valor and d2.codigo = " & DECODIFICADORA.DECODIFICADORA_INVENTARIO_ZONAS & _
                   filtro & _
                   " ORDER BY ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSINVENTARIO
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,TIPO_ID FROM INVENTARIO ORDER BY TIPO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function duplicar(ID As Long) As Long
    On Error GoTo fallo
        Me.Carga ID
        Me.setNOMBRE = Me.getNOMBRE & " (DUPLICADO)"
        Dim idNuevo As Long
        idNuevo = Me.Insertar
        duplicar = idNuevo
        Exit Function
fallo:
        duplicar = 0
End Function

