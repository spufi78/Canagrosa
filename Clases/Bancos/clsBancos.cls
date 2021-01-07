VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSBANCOS
'***************************************************************
Private ID_BANCO As Long
Private DESCRIPCION As String
Private ccc As String
Private SUBCUENTA As String
Private IMPORTE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSBANCOS
'***************************************************************
Public Property Let setID_BANCO(ByVal dato As Long)
        ID_BANCO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setCCC(ByVal dato As String)
        ccc = dato
End Property
Public Property Let setSUBCUENTA(ByVal dato As String)
        SUBCUENTA = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSBANCOS
'***************************************************************
Public Property Get getID_BANCO() As Long
        getID_BANCO = ID_BANCO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getCCC() As String
        getCCC = ccc
End Property
Public Property Get getSUBCUENTA() As String
        getSUBCUENTA = SUBCUENTA
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSBANCOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM BANCOS WHERE ID_BANCO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_BANCO = rs("ID_BANCO")
                DESCRIPCION = rs("DESCRIPCION")
                ccc = rs("CCC")
                SUBCUENTA = rs("SUBCUENTA")
                IMPORTE = rs("IMPORTE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsBancos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_BANCO) FROM BANCOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_BANCO = 1
        Else
                ID_BANCO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO BANCOS " & _
                   "  VALUES (" & _
                        ID_BANCO & "," & "'" & DESCRIPCION & "'" & "," & "'" & ccc & "','" & SUBCUENTA & "','0'" & _
                ")"
        execute_bd consulta
        Insertar = ID_BANCO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsBancos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE BANCOS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " SUBCUENTA = '" & SUBCUENTA & "'," & _
                        " CCC = '" & ccc & "'" & _
                " WHERE ID_BANCO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsBancos)", vbCritical, Err.Description
End Function
Public Function actualizarSaldo(ID_BANCO As Long, IMPORTE As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE BANCOS SET " & _
                        " IMPORTE = '" & IMPORTE & "'" & _
                " WHERE ID_BANCO = " & ID_BANCO
        execute_bd consulta
        actualizarSaldo = True
        Exit Function
fallo:
        actualizarSaldo = False
        MsgBox "Error al actualizarSaldo (clsBancos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM BANCOS " & _
                "    WHERE ID_BANCO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsBancos)", vbCritical, Err.Description
End Function
Public Function Listado(DESCRIPCION As String, ccc As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1 "
        If DESCRIPCION <> "" Then
            filtro = filtro & " AND A.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        If ccc <> "" Then
            filtro = filtro & " AND A.CCC LIKE '%" & ccc & "%'"
        End If
        consulta = "SELECT A.ID_BANCO,A.DESCRIPCION,A.CCC,A.SUBCUENTA,A.IMPORTE " & _
                   "  FROM BANCOS A " & _
                   filtro & _
                   " ORDER BY A.DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoId(ID As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE A.ID_BANCO = " & ID
        consulta = "SELECT A.ID_BANCO,A.DESCRIPCION,A.CCC,A.SUBCUENTA,A.IMPORTE " & _
                   "  FROM BANCOS A " & _
                   filtro & _
                   " ORDER BY A.DESCRIPCION"
        Set ListadoId = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSBANCOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_BANCO,DESCRIPCION FROM BANCOS ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setFILTRO = filtro
        .setTABLA = "BANCOS"
        .setDESCRIPCION = "Bancos"
        .setPK = "ID_BANCO"
        .setCAMPO = "DESCRIPCION"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function


