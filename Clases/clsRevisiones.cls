VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRevisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREVISIONES
'***************************************************************
Private ID_REVISION As Long
Private TOBJETO As Long
Private COBJETO As Long
Private USUARIO_ID As Long
Private FTIMESTP As String
Private OBSERVACIONES As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREVISIONES
'***************************************************************
Public Property Let setID_REVISION(ByVal dato As Long)
        ID_REVISION = dato
End Property
Public Property Let setTOBJETO(ByVal dato As Long)
        TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As Long)
        COBJETO = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREVISIONES
'***************************************************************
Public Property Get getID_REVISION() As Long
        getID_REVISION = ID_REVISION
End Property
Public Property Get getTOBJETO() As Long
        getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As Long
        getCOBJETO = COBJETO
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREVISIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REVISIONES WHERE ID_REVISION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_REVISION = rs("ID_REVISION")
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                USUARIO_ID = rs("USUARIO_ID")
                FTIMESTP = rs("FTIMESTP")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRevisiones)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_REVISION) FROM REVISIONES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_REVISION = 1
        Else
                ID_REVISION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO REVISIONES " & _
                   "  VALUES (" & _
                        ID_REVISION & "," & TOBJETO & "," & COBJETO & "," & _
                        USUARIO.getID_EMPLEADO & ",CURRENT_TIMESTAMP," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_REVISION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRevisiones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REVISIONES SET " & _
                        " TOBJETO = " & TOBJETO & "," & _
                        " COBJETO = " & COBJETO & "," & _
                        " USUARIO_ID = " & USUARIO.getID_EMPLEADO & "," & _
                        " FTIMESTP = CURRENT_TIMESTAMP " & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE ID_REVISION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRevisiones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM REVISIONES " & _
                "    WHERE ID_REVISION = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRevisiones)", vbCritical, Err.Description
End Function
Public Function Listado(TOBJETO As Long, COBJETO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_REVISION,A.FTIMESTP,U.NOMBRE,A.OBSERVACIONES " & _
                   "  FROM REVISIONES A " & _
                   "  LEFT JOIN usuarios U ON A.USUARIO_ID = U.ID_EMPLEADO " & _
                   " WHERE TOBJETO = " & TOBJETO & _
                   "   AND COBJETO = " & COBJETO & _
                   " ORDER BY FTIMESTP"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSREVISIONES
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_REVISION,TOBJETO FROM REVISIONES ORDER BY TOBJETO"
        Set Listado_Combo = datos_bd(consulta)
End Function
