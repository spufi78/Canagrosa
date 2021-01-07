VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCertificator_vida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCERTIFICATOR_VIDA
'***************************************************************
Private ID As Long
Private TOBJETO As Long
Private COBJETO As Long
Private ESTADO_ID As Long
Private EMPLEADO_ID As Long
Private TS As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSCERTIFICATOR_VIDA
'***************************************************************
Public Enum LISTA_CLSCERTIFICATOR_VIDA_COLS
        ENUM_ID = 1
        ENUM_TOBJETO = 2
        ENUM_COBJETO = 3
        ENUM_ESTADO_ID = 4
        ENUM_TS = 5
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCERTIFICATOR_VIDA
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setTOBJETO(ByVal dato As Long)
        TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As Long)
        COBJETO = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCERTIFICATOR_VIDA
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getTOBJETO() As Long
        getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As Long
        getCOBJETO = COBJETO
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCERTIFICATOR_VIDA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CERTIFICATOR_VIDA WHERE ID = " & ID
        Set rs = datos_bd_metrologia(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                ESTADO_ID = rs("ESTADO_ID")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                TS = rs("TS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCertificator_vida)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM CERTIFICATOR_VIDA"
        Set rs = datos_bd_metrologia(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CERTIFICATOR_VIDA " & _
                   "  VALUES (" & _
                        ID & "," & TOBJETO & "," & COBJETO & "," & _
                        ESTADO_ID & "," & EMPLEADO_ID & ",CURRENT_TIMESTAMP" & _
                ")"
        execute_bd_metrologia consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCertificator_vida)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CERTIFICATOR_VIDA SET " & _
                        " TOBJETO = " & TOBJETO & "," & _
                        " COBJETO = " & COBJETO & "," & _
                        " ESTADO_ID = " & ESTADO_ID & "," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " TS = '" & TS & "'" & _
                " WHERE ID = " & ID
        execute_bd_metrologia consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCertificator_vida)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CERTIFICATOR_VIDA " & _
                "    WHERE ID = " & ID
        execute_bd_metrologia consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCertificator_vida)", vbCritical, Err.Description
End Function
Public Function Listado(TOBJETO As Long, COBJETO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "select a.id,a.ts,CONCAT(b.NOMBRE,' ',b.APELLIDOS) AS NOMBRE,c.DESCRIPCION from certificator_vida a " & _
                    " left join geslab_canagrosa.usuarios b on a.EMPLEADO_ID = b.ID_EMPLEADO " & _
                    " left join geslab_canagrosa.decodificadora c on a.ESTADO_ID = c.VALOR and c.CODIGO = " & DECODIFICADORA.C_CERTIFICATOR_ESTADOS & _
                    " where tobjeto = " & TOBJETO & " and cobjeto = " & COBJETO & _
                    " order by ts desc"
        Set Listado = datos_bd_metrologia(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSCERTIFICATOR_VIDA
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd_metrologia(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,TOBJETO FROM CERTIFICATOR_VIDA ORDER BY TOBJETO"
        Set Listado_Combo = datos_bd_metrologia(consulta)
End Function
