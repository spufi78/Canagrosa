VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_req_familias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_REQ_FAMILIAS
'***************************************************************
Private ID_FAMILIA_REQ As Long
Private nombre As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_REQ_FAMILIAS
'***************************************************************
Public Property Let setID_FAMILIA_REQ(ByVal dato As Long)
        ID_FAMILIA_REQ = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_REQ_FAMILIAS
'***************************************************************
Public Property Get getID_FAMILIA_REQ() As Long
        getID_FAMILIA_REQ = ID_FAMILIA_REQ
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_REQ_FAMILIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_REQ_FAMILIAS WHERE ID_FAMILIA_REQ = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FAMILIA_REQ = rs("ID_FAMILIA_REQ")
                nombre = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_req_familias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_FAMILIA_REQ) FROM CA_REQ_FAMILIAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FAMILIA_REQ = 1
        Else
                ID_FAMILIA_REQ = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CA_REQ_FAMILIAS " & _
                   "  VALUES (" & _
                        ID_FAMILIA_REQ & "," & "'" & nombre & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_FAMILIA_REQ
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_req_familias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_REQ_FAMILIAS SET " & _
                        " NOMBRE = '" & nombre & "'" & _
                " WHERE ID_FAMILIA_REQ = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_req_familias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        execute_bd "DELETE FROM CA_REQ WHERE FAMILIA_REQ_ID = " & ID
        execute_bd "DELETE FROM CA_REQ_FAMILIAS WHERE ID_FAMILIA_REQ = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_req_familias)", vbCritical, Err.Description
End Function
Public Function Listado(nombre As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT a.ID_FAMILIA_REQ,a.NOMBRE, COUNT(distinct b.ID_REQ) " & _
                    "  FROM ca_req_familias a " & _
                    "  LEFT JOIN ca_req b on a.ID_FAMILIA_REQ = b.FAMILIA_REQ_ID " & _
                    " WHERE a.NOMBRE LIKE '%" & nombre & "%' " & _
                    "   AND a.ID_FAMILIA_REQ > 0 " & _
                    " GROUP BY 1 " & _
                    " ORDER BY a.NOMBRE "
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoId(idFamila As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT a.ID_FAMILIA_REQ,a.NOMBRE, COUNT(distinct b.ID_REQ) " & _
                    "  FROM ca_req_familias a " & _
                    "  LEFT JOIN ca_req b on a.ID_FAMILIA_REQ = b.FAMILIA_REQ_ID " & _
                    " WHERE a.ID_FAMILIA_REQ  = " & idFamila & _
                    " GROUP BY 1 "
        Set ListadoId = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSCA_REQ_FAMILIAS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_FAMILIA_REQ,NOMBRE FROM CA_REQ_FAMILIAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
