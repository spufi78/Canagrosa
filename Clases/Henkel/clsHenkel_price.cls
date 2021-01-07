VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHenkel_price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsHenkel_price
'***************************************************************
Private CODIGO As String
Private DIMENSION_ID As Integer
Private DIMENSION As String
Private PRECIO As String
Private PRIMER As Integer
Private COMENTARIOS As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE clsHenkel_price
'***************************************************************

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsHenkel_price
'***************************************************************
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDIMENSION_ID(ByVal dato As Integer)
        DIMENSION_ID = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setPRIMER(ByVal dato As Integer)
        PRIMER = dato
End Property
Public Property Let setCOMENTARIOS(ByVal dato As String)
        COMENTARIOS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsHenkel_price
'***************************************************************
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDIMENSION_ID() As Integer
        getDIMENSION_ID = DIMENSION_ID
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getPRIMER() As Integer
        getPRIMER = PRIMER
End Property

Public Property Get getCOMENTARIOS() As String
        getCOMENTARIOS = COMENTARIOS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsHenkel_price
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_PRICE WHERE CODIGO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CODIGO = rs("CODIGO")
                DIMENSION_ID = rs("DIMENSION_ID")
                DIMENSION = rs("DIMENSION")
                PRECIO = rs("PRECIO")
                PRIMER = rs("PRIMER")
                COMENTARIOS = rs("COMENTARIOS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsHenkel_price)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(CODIGO) FROM HENKEL_PRICE"
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
        Set rs = datos_bd("select * from HENKEL_PRICE where codigo = '" & CODIGO & "'")
        If rs.RecordCount > 0 Then
            MsgBox "El código ya existe, no se puede insertar.", vbCritical, App.Title
        Else
            consulta = "INSERT INTO HENKEL_PRICE " & _
                       "  VALUES (" & _
                            "'" & CODIGO & "'," & DIMENSION_ID & ",'" & DIMENSION & "','" & PRECIO & "'," & PRIMER & "," & _
                            "'" & COMENTARIOS & "'" & _
                    ")"
            execute_bd consulta
            Insertar = CODIGO
        End If
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsHenkel_price)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE HENKEL_PRICE SET " & _
                        " DIMENSION = '" & DIMENSION & "'," & _
                        " DIMENSION_ID = '" & DIMENSION_ID & "'," & _
                        " PRIMER = '" & PRIMER & "'," & _
                        " PRECIO = '" & PRECIO & "'," & _
                        " COMENTARIOS = '" & COMENTARIOS & "'" & _
                " WHERE CODIGO = '" & CODIGO & "'"
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsHenkel_price)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HENKEL_PRICE " & _
                "    WHERE CODIGO = '" & ID & "'"
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsHenkel_price)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HENKEL_PRICE ORDER BY CODIGO"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE clsHenkel_price
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT CODIGO,DIMENSION FROM HENKEL_PRICE ORDER BY DIMENSION"
        Set Listado_Combo = datos_bd(consulta)
End Function
