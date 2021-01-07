VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEq_parametros_tecnicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQ_PARAMETROS_TECNICOS
'***************************************************************
Private ID_PARAMETRO As Long
Private NOMBRE As String
Private NORMA_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQ_PARAMETROS_TECNICOS
'***************************************************************
Public Property Let setID_PARAMETRO(ByVal dato As Long)
        ID_PARAMETRO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQ_PARAMETROS_TECNICOS
'***************************************************************
Public Property Get getID_PARAMETRO() As Long
        getID_PARAMETRO = ID_PARAMETRO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_PARAMETROS_TECNICOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EQ_PARAMETROS_TECNICOS WHERE ID_PARAMETRO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PARAMETRO = rs("ID_PARAMETRO")
                NOMBRE = rs("NOMBRE")
                NORMA_ID = rs("NORMA_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEq_parametros_tecnicos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_PARAMETRO) FROM EQ_PARAMETROS_TECNICOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PARAMETRO = 1
        Else
                ID_PARAMETRO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EQ_PARAMETROS_TECNICOS " & _
                   "  VALUES (" & _
                        ID_PARAMETRO & "," & "'" & NOMBRE & "'" & "," & NORMA_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_PARAMETRO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEq_parametros_tecnicos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_PARAMETROS_TECNICOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " NORMA_ID = " & NORMA_ID & "" & _
                " WHERE ID_PARAMETRO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEq_parametros_tecnicos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EQ_PARAMETROS_TECNICOS " & _
                "    WHERE ID_PARAMETRO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEq_parametros_tecnicos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.ID_PARAMETRO, A.NOMBRE, CONCAT(B.CODIGO,' ',B.NOMBRE), A.NORMA_ID " & _
                   "  FROM EQ_PARAMETROS_TECNICOS A LEFT JOIN CA_NORMAS B ON A.NORMA_ID = B.ID_NORMA " & _
                   " ORDER BY A.NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_PARAMETRO,NOMBRE FROM EQ_PARAMETROS_TECNICOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
