VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHenkel_sets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSHENKEL_SETS
'***************************************************************
Private CODIGO As String
Private DIMENSION_ID As Long
Private DIMENSION As String
Private PRECIO_E As String
Private PRECIO_ABCD As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSHENKEL_SETS
'***************************************************************
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDIMENSION_ID(ByVal dato As Long)
        DIMENSION_ID = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setPRECIO_E(ByVal dato As String)
        PRECIO_E = dato
End Property
Public Property Let setPRECIO_ABCD(ByVal dato As String)
        PRECIO_ABCD = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSHENKEL_SETS
'***************************************************************
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDIMENSION_ID() As Long
        getDIMENSION_ID = DIMENSION_ID
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getPRECIO_E() As String
        getPRECIO_E = PRECIO_E
End Property
Public Property Get getPRECIO_ABCD() As String
        getPRECIO_ABCD = PRECIO_ABCD
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSHENKEL_SETS
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_SETS WHERE CODIGO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CODIGO = rs("CODIGO")
                DIMENSION_ID = rs("DIMENSION_ID")
                DIMENSION = rs("DIMENSION")
                PRECIO_E = rs("PRECIO_E")
                PRECIO_ABCD = rs("PRECIO_ABCD")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsHenkel_sets)", vbCritical, Err.Description
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO HENKEL_SETS " & _
                   "  VALUES (" & _
                        "'" & CODIGO & "'" & "," & DIMENSION_ID & ",'" & DIMENSION & "','" & PRECIO_E & "','" & _
                        PRECIO_ABCD & _
                "')"
        execute_bd consulta
        Insertar = CODIGO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsHenkel_sets)", vbCritical, Err.Description
End Function
Public Function Modificar(CODIGO As String, DIMENSION_ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE HENKEL_SETS SET " & _
                        " DIMENSION_ID = " & DIMENSION_ID & "," & _
                        " DIMENSION = '" & DIMENSION & "'," & _
                        " PRECIO_E = '" & PRECIO_E & "'," & _
                        " PRECIO_ABCD = '" & PRECIO_ABCD & "'" & _
                " WHERE CODIGO = '" & CODIGO & "' and DIMENSION_ID = " & DIMENSION_ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsHenkel_sets)", vbCritical, Err.Description
End Function
Public Function Eliminar(CODIGO As String, DIMENSION_ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HENKEL_SETS " & _
                "    WHERE CODIGO = " & CODIGO & _
                " and DIMENSION_ID = " & DIMENSION_ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsHenkel_sets)", vbCritical, Err.Description
End Function
Public Function Listado(CODIGO As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.CODIGO,A.DIMENSION,A.PRECIO_E,A.PRECIO_ABCD,A.DIMENSION_ID " & _
                   " FROM HENKEL_SETS A " & _
                   " LEFT JOIN DECODIFICADORA B ON A.DIMENSION_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_DIMENSIONES & _
                   " WHERE A.CODIGO = '" & CODIGO & "'" & _
                   " ORDER BY B.VALOR"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSHENKEL_SETS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

