VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCentros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCENTROS
'***************************************************************
Private ID_CENTRO As Long
Private nombre As String
Private DIRECCION1 As String
Private DIRECCION2 As String
Private CIF As String
Private TELEFONO As String
Private FAX As String
Private EMAIL As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCENTROS
'***************************************************************
Public Property Let setID_CENTRO(ByVal dato As Long)
        ID_CENTRO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setDIRECCION1(ByVal dato As String)
        DIRECCION1 = dato
End Property
Public Property Let setDIRECCION2(ByVal dato As String)
        DIRECCION2 = dato
End Property
Public Property Let setCIF(ByVal dato As String)
        CIF = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
Public Property Let setFAX(ByVal dato As String)
        FAX = dato
End Property
Public Property Let setEMAIL(ByVal dato As String)
        EMAIL = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCENTROS
'***************************************************************
Public Property Get getID_CENTRO() As Long
        getID_CENTRO = ID_CENTRO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getDIRECCION1() As String
        getDIRECCION1 = DIRECCION1
End Property
Public Property Get getDIRECCION2() As String
        getDIRECCION2 = DIRECCION2
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
Public Property Get getEMAIL() As String
        getEMAIL = EMAIL
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCENTROS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CENTROS WHERE ID_CENTRO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CENTRO = rs("ID_CENTRO")
                nombre = rs("NOMBRE")
                DIRECCION1 = rs("DIRECCION1")
                DIRECCION2 = rs("DIRECCION2")
                CIF = rs("CIF")
                TELEFONO = rs("TELEFONO")
                FAX = rs("FAX")
                EMAIL = rs("EMAIL")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCentros)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_CENTRO) FROM CENTROS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_CENTRO = 1
        Else
                ID_CENTRO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CENTROS " & _
                   "  VALUES (" & _
                        ID_CENTRO & "," & "'" & nombre & "'" & "," & "'" & DIRECCION1 & "'" & "," & _
                        "'" & DIRECCION2 & "'" & "," & "'" & CIF & "'" & "," & "'" & TELEFONO & "'" & "," & _
                        "'" & FAX & "'" & "," & "'" & EMAIL & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_CENTRO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCentros)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CENTROS SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " DIRECCION1 = '" & DIRECCION1 & "'," & _
                        " DIRECCION2 = '" & DIRECCION2 & "'," & _
                        " CIF = '" & CIF & "'," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " FAX = '" & FAX & "'," & _
                        " EMAIL = '" & EMAIL & "'" & _
                " WHERE ID_CENTRO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCentros)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CENTROS " & _
                "    WHERE ID_CENTRO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCentros)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CENTROS ORDER BY ID_CENTRO"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSCENTROS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_CENTRO,NOMBRE FROM CENTROS UNION SELECT 0,'' ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CENTROS"
        .setDESCRIPCION = "Centros"
        .setPK = "ID_CENTRO"
        .setFILTRO = filtro
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        .setFILTRO = ""
        Set .FORMULARIO = FORMULARIO
    End With
End Function
