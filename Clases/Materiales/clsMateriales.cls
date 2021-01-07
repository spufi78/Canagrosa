VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMateriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMATERIALES
'***************************************************************
Private ID_MATERIAL As Long
Private DEPARTAMENTO_ID As Long
Private TIPO_ID As Long
Private DESCRIPCION As String
Private RESPONSABLE_ID As Long
Private SERIE As String
Private ip As String
Private nombre As String
Private F_ALTA As String
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMATERIALES
'***************************************************************
Public Property Let setID_MATERIAL(ByVal dato As Long)
        ID_MATERIAL = dato
End Property
Public Property Let setDEPARTAMENTO_ID(ByVal dato As Long)
        DEPARTAMENTO_ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property
Public Property Let setSERIE(ByVal dato As String)
        SERIE = dato
End Property
Public Property Let setIP(ByVal dato As String)
        ip = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setF_ALTA(ByVal dato As String)
        F_ALTA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMATERIALES
'***************************************************************
Public Property Get getID_MATERIAL() As Long
        getID_MATERIAL = ID_MATERIAL
End Property
Public Property Get getDEPARTAMENTO_ID() As Long
        getDEPARTAMENTO_ID = DEPARTAMENTO_ID
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getSERIE() As String
        getSERIE = SERIE
End Property
Public Property Get getIP() As String
        getIP = ip
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getF_ALTA() As String
        getF_ALTA = F_ALTA
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMATERIALES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MATERIALES WHERE ID_MATERIAL = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_MATERIAL = rs("ID_MATERIAL")
                DEPARTAMENTO_ID = rs("DEPARTAMENTO_ID")
                TIPO_ID = rs("TIPO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                SERIE = rs("SERIE")
                ip = rs("IP")
                nombre = rs("NOMBRE")
                F_ALTA = rs("F_ALTA")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMateriales)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_MATERIAL) FROM MATERIALES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_MATERIAL = 1
        Else
                ID_MATERIAL = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO MATERIALES " & _
                   "  VALUES (" & _
                        ID_MATERIAL & "," & DEPARTAMENTO_ID & "," & TIPO_ID & "," & _
                        "'" & DESCRIPCION & "'" & "," & RESPONSABLE_ID & "," & "'" & SERIE & "'" & "," & _
                        "'" & ip & "'" & "," & "'" & nombre & "'" & "," & "'" & F_ALTA & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_MATERIAL
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMateriales)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MATERIALES SET " & _
                        " DEPARTAMENTO_ID = " & DEPARTAMENTO_ID & "," & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " RESPONSABLE_ID = " & RESPONSABLE_ID & "," & _
                        " SERIE = '" & SERIE & "'," & _
                        " IP = '" & ip & "'," & _
                        " NOMBRE = '" & nombre & "'," & _
                        " F_ALTA = '" & F_ALTA & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE ID_MATERIAL = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMateriales)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MATERIALES " & _
                "    WHERE ID_MATERIAL = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMateriales)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MATERIALES ORDER BY ID_MATERIAL"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMATERIALES
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_MATERIAL,DEPARTAMENTO_ID FROM MATERIALES ORDER BY DEPARTAMENTO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
