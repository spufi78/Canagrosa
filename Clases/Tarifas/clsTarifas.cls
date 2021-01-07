VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTARIFAS
'***************************************************************
Private ID_TARIFA As Long
Private NOMBRE As String
Private TARIFA_ORIGEN_ID As Long
Private EN_VIGOR As Integer
Private porcentaje As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTARIFAS
'***************************************************************
Public Property Let setID_TARIFA(ByVal dato As Long)
        ID_TARIFA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setTARIFA_ORIGEN_ID(ByVal dato As Long)
        TARIFA_ORIGEN_ID = dato
End Property
Public Property Let setEN_VIGOR(ByVal dato As Integer)
        EN_VIGOR = dato
End Property
Public Property Let setPORCENTAJE(ByVal dato As String)
        porcentaje = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTARIFAS
'***************************************************************
Public Property Get getID_TARIFA() As Long
        getID_TARIFA = ID_TARIFA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getTARIFA_ORIGEN_ID() As Long
        getTARIFA_ORIGEN_ID = TARIFA_ORIGEN_ID
End Property
Public Property Get getEN_VIGOR() As Integer
        getEN_VIGOR = EN_VIGOR
End Property
Public Property Get getPORCENTAJE() As String
        getPORCENTAJE = porcentaje
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTARIFAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM TARIFAS WHERE ID_TARIFA = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TARIFA = rs("ID_TARIFA")
                NOMBRE = rs("NOMBRE")
                TARIFA_ORIGEN_ID = rs("TARIFA_ORIGEN_ID")
                EN_VIGOR = rs("EN_VIGOR")
                porcentaje = rs("PORCENTAJE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTarifas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_TARIFA) FROM TARIFAS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_TARIFA = 1
        Else
                ID_TARIFA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO TARIFAS " & _
                   "  VALUES (" & _
                        ID_TARIFA & ",'" & NOMBRE & "'," & TARIFA_ORIGEN_ID & "," & _
                        EN_VIGOR & ",'" & Replace(porcentaje, ",", ".") & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_TARIFA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTarifas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE TARIFAS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_TARIFA = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTarifas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        Dim rs As ADODB.Recordset
        CONSULTA = "SELECT * FROM CLIENTES WHERE TARIFA_ID = " & ID & " AND ANULADO = 0 "
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede dar de baja la tarifa. Existen clientes que la tienen asignada.", vbExclamation, App.Title
            Eliminar = False
            Exit Function
        Else
            execute_bd "UPDATE TARIFAS SET EN_VIGOR = 0 WHERE ID_TARIFA = " & ID
'            execute_bd "DELETE FROM TARIFAS WHERE ID_TARIFA = " & ID
'            execute_bd "DELETE FROM TARIFAS_PRECIOS WHERE TARIFA_ID = " & ID
            Eliminar = True
        End If
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTarifas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.ID_TARIFA,A.NOMBRE, B.NOMBRE, A.PORCENTAJE, A.EN_VIGOR, COUNT(*) " & _
                   " FROM TARIFAS A " & _
                   " LEFT JOIN TARIFAS B ON A.TARIFA_ORIGEN_ID = B.ID_TARIFA " & _
                   " LEFT JOIN CLIENTES C ON A.ID_TARIFA = C.TARIFA_ID AND C.ANULADO = 0 " & _
                   " WHERE A.ID_TARIFA >= 0 " & _
                   " GROUP BY A.ID_TARIFA,A.NOMBRE,B.NOMBRE,A.PORCENTAJE, A.EN_VIGOR " & _
                   " ORDER BY A.NOMBRE"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Clientes(ID As Long) As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.ID_CLIENTE,A.NOMBRE " & _
                   " FROM CLIENTES A  " & _
                   " WHERE A.TARIFA_ID= " & ID & _
                   "   AND A.ANULADO = 0 " & _
                   " ORDER BY A.NOMBRE"
        Set Listado_Clientes = datos_bd(CONSULTA)
End Function
Public Function Listado_por_nombre() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.NOMBRE, B.NOMBRE, A.PORCENTAJE, A.ID_TARIFA " & _
                   " FROM TARIFAS A, TARIFAS B " & _
                   " WHERE A.ID_TARIFA >= 0 " & _
                   "   AND A.TARIFA_ORIGEN_ID = B.ID_TARIFA " & _
                   "   AND A.EN_VIGOR = 1 " & _
                   " ORDER BY A.NOMBRE"
        Set Listado_por_nombre = datos_bd(CONSULTA)
End Function
Public Function Listado_por_nombre_Cliente(cliente As Long) As ADODB.Recordset
        Dim CONSULTA As String
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cliente
        CONSULTA = "SELECT A.NOMBRE, B.NOMBRE, A.PORCENTAJE, A.ID_TARIFA " & _
                   " FROM TARIFAS A, TARIFAS B " & _
                   " WHERE A.ID_TARIFA = " & oCliente.getTARIFA_ID & _
                   "   AND A.ID_TARIFA > 0 " & _
                   "   AND A.EN_VIGOR = 1 " & _
                   "   AND A.TARIFA_ORIGEN_ID = B.ID_TARIFA " & _
                   " ORDER BY A.NOMBRE"
        Set Listado_por_nombre_Cliente = datos_bd(CONSULTA)
End Function

Public Function Listado_Combo() As ADODB.Recordset

        Dim CONSULTA As String
        CONSULTA = "SELECT ID_TARIFA,NOMBRE FROM TARIFAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TARIFAS"
        .setDESCRIPCION = "Tarifas"
        .setPK = "ID_TARIFA"
        .setCAMPO = "NOMBRE"
        .setFILTRO = " EN_VIGOR = 1 " & filtro
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Modificar_Porcentaje(TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE TARIFAS " & _
                   "   SET PORCENTAJE = PORCENTAJE + " & CSng(porcentaje) & _
                   " WHERE ID_TARIFA = " & TARIFA
        execute_bd CONSULTA
        Modificar_Porcentaje = True
        Exit Function
fallo:
        Modificar_Porcentaje = False
        MsgBox "Error al Modificar el porcentaje (clsTarifas)", vbCritical, Err.Description
End Function
