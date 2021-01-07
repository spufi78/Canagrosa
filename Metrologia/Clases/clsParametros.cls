VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPARAMETROS
'***************************************************************
Private ID_PARAMETRO As Long
Private descripcion As String
Private VALOR As String
Private USUARIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPARAMETROS
'***************************************************************
Public Property Let setID_PARAMETRO(ByVal dato As Long)
        ID_PARAMETRO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        descripcion = dato
End Property
Public Property Let setVALOR(ByVal dato As String)
        VALOR = dato
End Property
Public Property Let setUSUARIO(ByVal dato As String)
        USUARIO = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPARAMETROS
'***************************************************************
Public Property Get getID_PARAMETRO() As Long
        getID_PARAMETRO = ID_PARAMETRO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = descripcion
End Property
Public Property Get getVALOR() As String
        getVALOR = VALOR
End Property
Public Property Get getUSUARIO() As String
        getUSUARIO = USUARIO
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPARAMETROS
'***************************************************************
Public Function Carga(PARAMETRO As Long, USUARIO As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PARAMETROS WHERE ID_PARAMETRO = " & PARAMETRO & " AND USUARIO = '" & USUARIO & "'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PARAMETRO = rs("ID_PARAMETRO")
                descripcion = rs("DESCRIPCION")
                VALOR = rs("VALOR")
                USUARIO = rs("USUARIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsParametros)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PARAMETRO) FROM PARAMETROS"
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
'        Me.CrearID
        consulta = "INSERT INTO PARAMETROS " & _
                   "  VALUES (" & _
                        ID_PARAMETRO & "," & "'" & descripcion & "'" & "," & "'" & VALOR & "','" & USUARIO & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_PARAMETRO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsParametros)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, cUSUARIO As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PARAMETROS SET " & _
                        " DESCRIPCION = '" & descripcion & "'," & _
                        " USUARIO = '" & USUARIO & "'," & _
                        " VALOR = '" & VALOR & "'" & _
                " WHERE ID_PARAMETRO = " & ID & _
                "   AND USUARIO = '" & cUSUARIO & "'"
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsParametros)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, cUSUARIO As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PARAMETROS " & _
                "    WHERE ID_PARAMETRO = " & ID & _
                "      AND USUARIO = '" & cUSUARIO & "'"
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsParametros)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PARAMETROS ORDER BY ID_PARAMETRO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PARAMETRO,DESCRIPCION FROM PARAMETROS ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function actualizar_valor(PARAMETRO As Long, cUSUARIO As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PARAMETROS SET " & _
                        " VALOR = '" & VALOR & "'" & _
                " WHERE ID_PARAMETRO = " & PARAMETRO & _
                "   AND USUARIO = '" & cUSUARIO & "'"
        execute_bd consulta
        actualizar_valor = True
        Exit Function
fallo:
        actualizar_valor = False
        MsgBox "Error al actualizar_valor (clsParametros)", vbCritical, Err.Description
End Function


