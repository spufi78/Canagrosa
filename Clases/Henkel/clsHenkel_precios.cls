VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHenkel_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSHENKEL_PRECIOS
'***************************************************************
Private CODIGO As String
Private PRECIO_E As String
Private PRECIO_ABCD As String
Private COMENTARIOS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSHENKEL_PRECIOS
'***************************************************************
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setPRECIO_E(ByVal dato As String)
        PRECIO_E = dato
End Property
Public Property Let setPRECIO_ABCD(ByVal dato As String)
        PRECIO_ABCD = dato
End Property
Public Property Let setCOMENTARIOS(ByVal dato As String)
        COMENTARIOS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSHENKEL_PRECIOS
'***************************************************************
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getPRECIO_E() As String
        getPRECIO_E = PRECIO_E
End Property
Public Property Get getPRECIO_ABCD() As String
        getPRECIO_ABCD = PRECIO_ABCD
End Property
Public Property Get getCOMENTARIOS() As String
        getCOMENTARIOS = COMENTARIOS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSHENKEL_PRECIOS
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_PRECIOS WHERE CODIGO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CODIGO = rs("CODIGO")
                PRECIO_E = rs("PRECIO_E")
                PRECIO_ABCD = rs("PRECIO_ABCD")
                COMENTARIOS = rs("COMENTARIOS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsHenkel_precios)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(CODIGO) FROM HENKEL_PRECIOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                CODIGO = 1
        Else
                CODIGO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("select * from henkel_precios where codigo = '" & CODIGO & "'")
        If rs.RecordCount > 0 Then
            MsgBox "El código ya existe, no se puede insertar.", vbCritical, App.Title
        Else
            consulta = "INSERT INTO HENKEL_PRECIOS " & _
                       "  VALUES (" & _
                            "'" & CODIGO & "'" & ",'" & PRECIO_E & "','" & PRECIO_ABCD & "'," & _
                            "'" & COMENTARIOS & "'" & _
                    ")"
            execute_bd consulta
            Insertar = CODIGO
        End If
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsHenkel_precios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE HENKEL_PRECIOS SET " & _
                        " PRECIO_E = '" & PRECIO_E & "'," & _
                        " PRECIO_ABCD = '" & PRECIO_ABCD & "'," & _
                        " COMENTARIOS = '" & COMENTARIOS & "'" & _
                " WHERE CODIGO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsHenkel_precios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HENKEL_PRECIOS " & _
                "    WHERE CODIGO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsHenkel_precios)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_PRECIOS ORDER BY CODIGO"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSHENKEL_PRECIOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT CODIGO,PRECIO_E FROM HENKEL_PRECIOS ORDER BY PRECIO_E"
        Set Listado_Combo = datos_bd(consulta)
End Function
