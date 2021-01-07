VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifas_codigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTARIFAS_CODIGOS
'***************************************************************
Private ID_CODIGO As Long
Private CODIGO As String
Private DESCRIPCION As String
Private FAMILIA_CODIGO_ID As String
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTARIFAS_CODIGOS
'***************************************************************
Public Property Let setID_CODIGO(ByVal dato As Long)
        ID_CODIGO = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFAMILIA_CODIGO_ID(ByVal dato As String)
        FAMILIA_CODIGO_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTARIFAS_CODIGOS
'***************************************************************
Public Property Get getID_CODIGO() As Long
        getID_CODIGO = ID_CODIGO
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFAMILIA_CODIGO_ID() As String
        getFAMILIA_CODIGO_ID = FAMILIA_CODIGO_ID
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTARIFAS_CODIGOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM TARIFAS_CODIGOS WHERE ID_CODIGO = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CODIGO = rs("ID_CODIGO")
                CODIGO = rs("CODIGO")
                DESCRIPCION = rs("DESCRIPCION")
                FAMILIA_CODIGO_ID = rs("FAMILIA_CODIGO_ID")
                PRECIO = rs("PRECIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTarifas_codigos)", vbCritical, Err.Description
End Function
Public Function Codigo_Bano(BANO As Long) As String
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.CODIGO " & _
                   "  FROM TARIFAS_CODIGOS A, BANOS B " & _
                   " WHERE A.ID_CODIGO = B.TARIFA_CODIGO_ID " & _
                   "   AND B.ID_BANO = " & BANO
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
           Codigo_Bano = ""
        Else
           Codigo_Bano = rs(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Codigo_Bano = ""
        MsgBox "Error al cargar los datos(clsTarifas_codigos)", vbCritical, Err.Description
End Function
Public Function Codigo_TipoAnalisis(TA As Long) As String
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.CODIGO " & _
                   "  FROM TARIFAS_CODIGOS A, TIPOS_ANALISIS B " & _
                   " WHERE A.ID_CODIGO = B.TARIFA_CODIGO_ID " & _
                   "   AND B.ID_TIPO_ANALISIS = " & TA
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
           Codigo_TipoAnalisis = ""
        Else
           Codigo_TipoAnalisis = rs(0)
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Codigo_TipoAnalisis = ""
        MsgBox "Error al cargar los datos(clsTarifas_codigos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_CODIGO) FROM TARIFAS_CODIGOS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_CODIGO = 1
        Else
                ID_CODIGO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO TARIFAS_CODIGOS " & _
                   "  VALUES (" & _
                        ID_CODIGO & "," & "'" & CODIGO & "'" & "," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & FAMILIA_CODIGO_ID & "'" & ",'" & PRECIO & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_CODIGO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTarifas_codigos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE TARIFAS_CODIGOS SET " & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FAMILIA_CODIGO_ID = '" & FAMILIA_CODIGO_ID & "'," & _
                        " PRECIO = '" & PRECIO & "'" & _
                " WHERE ID_CODIGO = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTarifas_codigos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM TARIFAS_CODIGOS " & _
                "    WHERE ID_CODIGO = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTarifas_codigos)", vbCritical, Err.Description
End Function
Public Function Listado(desc As String, cod As String, familia As Integer) As ADODB.Recordset
        Dim CONSULTA As String
        Dim s As String
        s = ""
        If Trim(desc) <> "" Then
            s = s & " AND TC.DESCRIPCION LIKE '%" & desc & "%'"
        End If
        If Trim(cod) <> "" Then
            s = s & " AND TC.CODIGO LIKE '%" & cod & "%'"
        End If
        If familia <> 0 Then
            s = s & " AND TC.FAMILIA_CODIGO_ID = " & familia
        End If
        CONSULTA = "SELECT TC.ID_CODIGO,TC.CODIGO,TC.DESCRIPCION,TCF.DESCRIPCION,TC.FAMILIA_CODIGO_ID " & _
                   " FROM TARIFAS_CODIGOS TC,TARIFAS_CODIGOS_FAMILIAS TCF " & _
                   " WHERE TC.FAMILIA_CODIGO_ID = TCF.ID_FAMILIA_CODIGOS " & _
                   s & _
                   " ORDER BY CODIGO"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Filtro(familia As Integer, CODIGO As String, DESCRIPCION As String) As ADODB.Recordset
        Dim CONSULTA As String
        If familia <> 0 Then
            Dim sFamilia As String
            sFamilia = " AND TC.FAMILIA_CODIGO_ID = " & familia
        End If
        CONSULTA = "SELECT DISTINCT TC.ID_CODIGO,TC.CODIGO,TC.DESCRIPCION,TCF.DESCRIPCION,TC.FAMILIA_CODIGO_ID " & _
                   " FROM TARIFAS_CODIGOS TC,TARIFAS_CODIGOS_FAMILIAS TCF " & _
                   " WHERE TC.FAMILIA_CODIGO_ID = TCF.ID_FAMILIA_CODIGOS " & _
                   "   AND TC.CODIGO LIKE '%" & CODIGO & "%'" & _
                   "   AND TC.DESCRIPCION LIKE '%" & DESCRIPCION & "%'" & _
                   sFamilia & _
                   " ORDER BY CODIGO"
        Set Listado_Filtro = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_CODIGO,CODIGO FROM TARIFAS_CODIGOS ORDER BY CODIGO"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function Modificar_Precio(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE TARIFAS_CODIGOS SET " & _
                        " PRECIO = '" & PRECIO & "'" & _
                   " WHERE ID_CODIGO = " & ID
        execute_bd CONSULTA
        Modificar_Precio = True
        Exit Function
fallo:
        Modificar_Precio = False
        MsgBox "Error al Modificar el precio del código(clsTarifas_codigos)", vbCritical, Err.Description
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TARIFAS_CODIGOS"
        .setDESCRIPCION = "Códigos Tarifarios"
        .setPK = "ID_CODIGO"
        .setCAMPO = "concat('(',CODIGO,') ',DESCRIPCION)"
'        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
